using System.Text;
using System.Text.Json;
using tembu_server.Models;

namespace tembu_server.Services;

// Holt Mail-/Kalender-Batches ueber GraphDataService und schickt sie durch das konfigurierte
// LLM (AiService), um Entitaeten (Personen/Organisationen/Themen) und Beziehungen zwischen
// ihnen zu extrahieren — der Rohstoff fuer den Knowledge Graph.
//
// Welches LLM tatsaechlich antwortet (lokal via Ollama oder ein externer Provider), entscheidet
// die AI-Settings-Konfiguration des Kunden (ai-settings.json) — dieser Service ist providerneutral.
public class GraphExtractionService(GraphDataService graphData, AiService ai, ILogger<GraphExtractionService> logger)
{
    // Begrenzt die Batchgroesse pro LLM-Call, damit der Prompt im Kontextfenster kleinerer
    // lokaler Modelle bleibt und ein einzelner Parse-Fehler nicht den ganzen Sync verwirft.
    private const int BatchSize = 25;

    // from/to wirken nur auf einen Erstlauf (kein deltaLink) — grenzen ein, wie weit der
    // initiale Sync in die Historie zurückgeht, statt die ganze Mailbox/den ganzen Kalender
    // zu verarbeiten.
    public async Task<GraphExtractionResult> ExtractAsync(string bearerToken, string? mailDeltaLink = null, string? calendarDeltaLink = null, DateTimeOffset? from = null, DateTimeOffset? to = null, IReadOnlyCollection<string>? excludeFolderIds = null)
    {
        var mail   = await graphData.GetMailDeltaAsync(bearerToken, mailDeltaLink, from: from, to: to, excludeFolderIds: excludeFolderIds);
        var events = await graphData.GetCalendarDeltaAsync(bearerToken, calendarDeltaLink, from: from, to: to);

        // Dictionary statt Liste, damit dieselbe Person/Organisation ueber mehrere Batches
        // (Mail + Kalender) hinweg auf denselben Knoten gemerged wird statt Duplikate zu erzeugen.
        var entities  = new Dictionary<string, GraphEntity>(StringComparer.OrdinalIgnoreCase);
        var relations = new List<GraphRelation>();

        foreach (var batch in mail.Items.Chunk(BatchSize))
            await ExtractBatchAsync(PromptBuilder.BuildMailExtraction(batch), PromptBuilder.MailExtractionSchema, PromptBuilder.MailRelationTypes, BuildNameEmailMap(batch), entities, relations);

        foreach (var batch in events.Items.Chunk(BatchSize))
            await ExtractBatchAsync(PromptBuilder.BuildEventExtraction(batch), PromptBuilder.EventExtractionSchema, PromptBuilder.EventRelationTypes, BuildNameEmailMap(batch), entities, relations);

        return new GraphExtractionResult
        {
            Entities          = entities.Values.ToList(),
            Relations         = relations,
            MailDeltaLink     = mail.DeltaLink,
            CalendarDeltaLink = events.DeltaLink,
        };
    }

    // Best-effort Mapping fuer haeufig vom Modell erfundene Synonyme auf die erlaubten Typen,
    // bevor eine Relation als nicht zuordenbar verworfen wird (aus realen Testlaeufen von
    // qwen2.5:14b beobachtet — bei anderen Modellen ggf. erweitern).
    private static readonly Dictionary<string, string> RelationTypeSynonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        ["communication"]            = "communicated_with",
        ["sent-message"]             = "communicated_with",
        ["sent_message"]             = "communicated_with",
        ["responded-to"]             = "communicated_with",
        ["responded_to"]             = "communicated_with",
        ["requests_from"]            = "communicated_with",
        ["discusses"]                = "works_on",
        ["topic discussion"]         = "works_on",
        ["topic_discussion"]         = "works_on",
        ["file reference"]           = "works_on",
        ["file_reference"]           = "works_on",
        ["works_for"]                = "belongs_to",
        ["organization affiliation"] = "belongs_to",
        ["organization_affiliation"] = "belongs_to",
    };

    private static Dictionary<string, string> BuildNameEmailMap(IEnumerable<GraphMailItem> items)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in items)
        {
            AddToNameMap(map, m.FromName, m.From);
            for (var i = 0; i < m.ToRecipients.Count && i < m.ToRecipientNames.Count; i++)
                AddToNameMap(map, m.ToRecipientNames[i], m.ToRecipients[i]);
        }
        return map;
    }

    private static Dictionary<string, string> BuildNameEmailMap(IEnumerable<GraphEventItem> items)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in items)
            for (var i = 0; i < e.Attendees.Count && i < e.AttendeeNames.Count; i++)
                AddToNameMap(map, e.AttendeeNames[i], e.Attendees[i]);
        return map;
    }

    private static void AddToNameMap(Dictionary<string, string> map, string name, string email)
    {
        if (name.Length > 0 && email.Length > 0) map[name] = email;
    }

    private async Task ExtractBatchAsync(string prompt, string schema, IReadOnlyCollection<string> allowedRelationTypes, IReadOnlyDictionary<string, string> nameMap, Dictionary<string, GraphEntity> entities, List<GraphRelation> relations)
    {
        string raw;
        try
        {
            raw = await ai.CallAsync(prompt, schema);
        }
        catch (Exception ex)
        {
            // Ein einzelner Batch-Fehler (z.B. Ollama kurz nicht erreichbar) soll nicht den
            // gesamten Sync abbrechen — der Batch wird beim naechsten Lauf erneut versucht,
            // da der deltaLink erst nach erfolgreichem Abschluss weitergereicht wird.
            logger.LogWarning(ex, "GraphExtraction: LLM-Aufruf fehlgeschlagen, Batch wird uebersprungen");
            return;
        }

        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(StripMarkdownFence(raw));
        }
        catch (JsonException)
        {
            logger.LogWarning("GraphExtraction: LLM-Antwort war kein valides JSON, Batch wird uebersprungen: {Raw}", raw.Length > 300 ? raw[..300] : raw);
            return;
        }

        using (doc)
        {
            // Ids, die das Modell trotz Vorgabe/Schema nicht als E-Mail bzw. Slug ausgibt (z.B.
            // reine Zahlen), werden ueber die tatsaechliche Absender-/Empfaenger-Adresse aus den
            // Rohdaten (nameMap) korrigiert statt dem Modell zu vertrauen. idRemap haelt fest,
            // welche vom Modell erfundene Id sich geaendert hat, damit die zugehoerigen Relations
            // in diesem Batch konsistent mitziehen.
            var idRemap = new Dictionary<string, string>();

            if (doc.RootElement.TryGetProperty("entities", out var ents))
                foreach (var e in ents.EnumerateArray())
                {
                    var rawId = GetString(e, "id");
                    var type  = GetString(e, "type");
                    var name  = GetString(e, "name");
                    var email = GetString(e, "email");
                    if (rawId.Length == 0) continue;

                    var id = type switch
                    {
                        "person" => ResolvePersonId(rawId, email, name, nameMap),
                        "organization" or "topic" => ResolveSlugId(rawId, name),
                        _ => rawId,
                    };
                    if (id != rawId) idRemap[rawId] = id;

                    entities[id] = new GraphEntity
                    {
                        Id    = id,
                        Type  = type,
                        Name  = name,
                        Email = type == "person" ? id : email,
                    };
                }

            if (doc.RootElement.TryGetProperty("relations", out var rels))
            {
                var droppedTypes = new List<string>();
                foreach (var r in rels.EnumerateArray())
                {
                    var from = GetString(r, "from");
                    var to   = GetString(r, "to");
                    var type = GetString(r, "type");
                    if (from.Length == 0 || to.Length == 0) continue;

                    if (!allowedRelationTypes.Contains(type))
                    {
                        if (RelationTypeSynonyms.TryGetValue(type, out var mapped)) type = mapped;
                        else { droppedTypes.Add(type); continue; }
                    }

                    relations.Add(new GraphRelation
                    {
                        From    = idRemap.GetValueOrDefault(from, from),
                        To      = idRemap.GetValueOrDefault(to, to),
                        Type    = type,
                        Context = GetString(r, "context"),
                        Date    = GetString(r, "date"),
                    });
                }

                if (droppedTypes.Count > 0)
                    logger.LogInformation("GraphExtraction: {Count} Relations mit nicht zuordenbarem Typ verworfen: {Types}",
                        droppedTypes.Count, string.Join(", ", droppedTypes.Distinct()));
            }
        }
    }

    private static string ResolvePersonId(string modelId, string email, string name, IReadOnlyDictionary<string, string> nameMap)
    {
        if (IsValidEmail(modelId)) return modelId.ToLowerInvariant();
        if (IsValidEmail(email)) return email.ToLowerInvariant();
        if (name.Length > 0 && nameMap.TryGetValue(name, out var resolved)) return resolved.ToLowerInvariant();
        return modelId; // nichts Besseres gefunden — Original-Id (z.B. eine vom Modell erfundene Zahl) beibehalten
    }

    private static bool IsValidEmail(string s)
    {
        var at = s.IndexOf('@');
        return at > 0 && s.IndexOf('.', at) > at + 1;
    }

    // organization/topic: das Modell vergibt trotz Vorgabe teils rein numerische Ids ("3", "7")
    // statt eines lesbaren Slugs — dann aus dem Namen selbst einen Slug generieren.
    private static string ResolveSlugId(string modelId, string name)
    {
        if (IsValidSlug(modelId)) return modelId;
        var slug = Slugify(name);
        return slug.Length > 0 ? slug : modelId;
    }

    private static bool IsValidSlug(string s) => s.Length > 0 && !s.All(char.IsDigit);

    private static string Slugify(string name)
    {
        var sb = new StringBuilder();
        var lastWasDash = true; // verhindert einen fuehrenden Bindestrich
        foreach (var c in name.ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(c)) { sb.Append(c); lastWasDash = false; }
            else if (!lastWasDash) { sb.Append('-'); lastWasDash = true; }
        }
        return sb.ToString().TrimEnd('-');
    }

    // Manche lokalen Modelle halten sich trotz Anweisung nicht an "kein Markdown" und wickeln
    // die Antwort in ```json ... ``` — defensiv entfernen statt den Batch deswegen zu verwerfen.
    private static string StripMarkdownFence(string s)
    {
        s = s.Trim();
        if (!s.StartsWith("```")) return s;
        var firstNewline = s.IndexOf('\n');
        s = firstNewline >= 0 ? s[(firstNewline + 1)..] : s;
        return s.EndsWith("```") ? s[..^3].Trim() : s.Trim();
    }

    private static string GetString(JsonElement el, string prop) =>
        el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
}
