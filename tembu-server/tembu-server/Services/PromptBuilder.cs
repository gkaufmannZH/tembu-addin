using tembu_server.Models;

namespace tembu_server.Services;

public static class PromptBuilder
{
    // Enum-Werte MÜSSEN Englisch bleiben — das Client-Rendering (detail.js/company.js)
    // vergleicht z.B. a.sentiment === 'positive' bzw. t.status === 'open' direkt.
    // Nur die Freitext-Felder (summary/nextStep/...) sollen in der Nutzersprache sein.
    private static string LanguageInstruction(string? lang) => lang switch
    {
        "en" => "Respond in English.",
        "fr" => "Réponds en français.",
        "es" => "Responde en español.",
        _    => "Antworte auf Deutsch.",
    };

    public static string Build(ContactData data)
    {
        var today = DateTime.Now.ToString("dd.MM.yyyy");

        var eLines = string.Join("\n", data.Emails.Take(60).Select(e =>
        {
            var dir = e.Direction == "received" ? "VON" : "AN";
            return "[" + e.DateStr + "] EMAIL " + dir + ": \"" + e.Subject + "\" - " + e.Preview;
        }));

        var mLines = string.Join("\n", data.Meetings.Take(40).Select(m =>
            "[" + m.DateStr + "] MEETING (" + m.DurationMin + "min): \"" + m.Subject + "\""));

        var rLines = string.Join("\n", data.Rumbles.Select(r =>
            "[" + r.DateStr + "] RUMBLE: \"" + r.Subject + "\""));

        var blocks = "";
        if (eLines.Length > 0) blocks += "E-MAILS:\n" + eLines + "\n\n";
        if (mLines.Length > 0) blocks += "MEETINGS:\n" + mLines + "\n\n";
        if (rLines.Length > 0) blocks += "RUMBLES:\n" + rLines + "\n\n";
        if (blocks.Length == 0) blocks = "Noch keine Interaktionen.\n\n";

        var jsonTemplate =
            "{\n" +
            "  \"summary\": \"2-3 Saetze zur Beziehung, Haeufigkeit, Ton\",\n" +
            "  \"sentiment\": \"positive|neutral|negative\",\n" +
            "  \"openPoints\": [\"Offener Punkt 1\"],\n" +
            "  \"themes\": [\n" +
            "    { \"name\": \"Thema\", \"status\": \"open|done\", \"summary\": \"Kurzbeschreibung\",\n" +
            "      \"interactions\": [{\"date\":\"YYYY-MM-DD\",\"type\":\"email|meeting\",\"subject\":\"Betreff\"}] }\n" +
            "  ],\n" +
            "  \"nextStep\": \"Konkrete Empfehlung fuer naechstes Gespraech\",\n" +
            "  \"background\": \"Oeffentlich bekannte Infos zu " + data.ContactName + ". Falls unbekannt: leer lassen.\"\n" +
            "}";

        return
            "Du bist ein persoenlicher Business-Assistent. Heute ist " + today + ". " + LanguageInstruction(data.Lang) + "\n" +
            "Analysiere alle Interaktionen mit \"" + data.ContactName + "\" (" + data.ContactEmail + ").\n\n" +
            blocks +
            "Antworte NUR mit validem JSON (kein Markdown).\n" +
            "Wichtig: Im interactions-Array jedes Themas ALLE zugehoerigen Interaktionen auflisten.\n" +
            jsonTemplate;
    }

    public static string BuildCompany(CompanyData data)
    {
        var today = DateTime.Now.ToString("dd.MM.yyyy");
        var contactNames = data.Contacts.Count > 0
            ? string.Join(", ", data.Contacts.Select(c => c.Name).Where(n => n.Length > 0))
            : "unbekannte Kontakte";

        var eLines = string.Join("\n", data.Emails.Take(80).Select(e =>
        {
            var dir = e.Direction == "received" ? "VON" : "AN";
            return "[" + e.DateStr + "] EMAIL " + dir + " " + e.Contact + ": \"" + e.Subject + "\"";
        }));

        var mLines = string.Join("\n", data.Meetings.Take(40).Select(m =>
            "[" + m.DateStr + "] MEETING mit " + m.Contact + " (" + m.DurationMin + "min): \"" + m.Subject + "\""));

        var blocks = "";
        if (eLines.Length > 0) blocks += "E-MAILS:\n" + eLines + "\n\n";
        if (mLines.Length > 0) blocks += "MEETINGS:\n" + mLines + "\n\n";
        if (blocks.Length == 0) blocks = "Noch keine Interaktionen gefunden.\n\n";

        var jsonTemplate =
            "{\n" +
            "  \"summary\": \"2-3 Saetze zur Gesamtbeziehung mit der Firma\",\n" +
            "  \"sentiment\": \"positive|neutral|negative\",\n" +
            "  \"openPoints\": [\"Offener Punkt 1\"],\n" +
            "  \"themes\": [\n" +
            "    { \"name\": \"Thema\", \"status\": \"open|done\", \"summary\": \"Kurzbeschreibung\", \"contacts\": [\"Name1\",\"Name2\"],\n" +
            "      \"interactions\": [{\"date\":\"YYYY-MM-DD\",\"type\":\"email|meeting\",\"contact\":\"Name\",\"subject\":\"Betreff\"}] }\n" +
            "  ],\n" +
            "  \"nextStep\": \"Konkrete Empfehlung fuer naechsten Schritt mit dieser Firma\"\n" +
            "}";

        return
            "Du bist ein Business-Assistent. Heute ist " + today + ". " + LanguageInstruction(data.Lang) + "\n" +
            "Analysiere meine Geschaeftsbeziehung mit der Firma \"" + data.CompanyName + "\" (Domain: " + data.Domain + ").\n" +
            "Bekannte Kontakte: " + contactNames + ".\n\n" +
            blocks +
            "Antworte NUR mit validem JSON (kein Markdown).\n" +
            "Wichtig: Im interactions-Array ALLE zugehoerigen Interaktionen auflisten, keine Auswahl.\n" +
            jsonTemplate;
    }

    // Begrenzt auf 40 Unterhaltungen/15 Mails pro Aufruf — der Client schickt ohnehin nur
    // die Unterhaltungen mit, die seit der letzten gespeicherten Zusammenfassung neue Mails
    // haben; diese Kappung ist nur ein zusaetzliches serverseitiges Sicherheitsnetz.
    public static string BuildConversationSummaries(ConversationSummaryRequest data)
    {
        var today = DateTime.Now.ToString("dd.MM.yyyy");

        var convBlocks = string.Join("\n\n", data.Conversations.Take(40).Select(c =>
        {
            var lines = string.Join("\n", c.Emails.Take(15).Select(e =>
            {
                var dir = e.Direction == "received" ? "VON" : "AN";
                return "[" + e.DateStr + "] " + dir + ": \"" + e.Subject + "\" - " + e.Preview;
            }));
            return "UNTERHALTUNG \"" + c.Id + "\" (Betreff: \"" + c.Subject + "\"):\n" + lines;
        }));

        var jsonTemplate =
            "{\n" +
            "  \"summaries\": [\n" +
            "    { \"id\": \"exakt die Unterhaltungs-Id von oben\", \"summary\": \"1-2 knappe Saetze: worum geht es, aktueller Stand\" }\n" +
            "  ]\n" +
            "}";

        return
            "Du bist ein persoenlicher Business-Assistent. Heute ist " + today + ". " + LanguageInstruction(data.Lang) + "\n" +
            "Fasse jede der folgenden E-Mail-Unterhaltungen mit \"" + data.ContactName + "\" (" + data.ContactEmail + ") in 1-2 knappen Saetzen zusammen.\n\n" +
            convBlocks + "\n\n" +
            "Antworte NUR mit validem JSON (kein Markdown). Fuer JEDE Unterhaltung von oben genau einen Eintrag mit exakt derselben id:\n" +
            jsonTemplate;
    }

    private const string GraphExtractionFormat =
        "{\n" +
        "  \"entities\": [{\"id\":\"...\",\"type\":\"person|organization|topic\",\"name\":\"...\",\"email\":\"...\"}],\n" +
        "  \"relations\": [{\"from\":\"id\",\"to\":\"id\",\"type\":\"...\",\"context\":\"kurzer Kontext\",\"date\":\"YYYY-MM-DD\"}]\n" +
        "}";

    // Ohne diese expliziten Verbote weicht qwen2.5:14b in der Praxis regelmaessig ab:
    // vergibt Zahlen statt E-Mail/Slug als id (bricht das Merging ueber Batches per
    // Dictionary-Key in GraphExtractionService), kombiniert Typen ("person|organization")
    // oder erfindet eigene relations.type-Werte statt der vorgegebenen Liste.
    private const string GraphExtractionRules =
        "REGELN (unbedingt einhalten):\n" +
        "- entities[].id bei type \"person\": IMMER die vollstaendige E-Mail-Adresse (z.B. \"anna.berger@zulieferer.ch\"). NIEMALS eine Zahl oder nur der Name.\n" +
        "- entities[].id bei type \"organization\"/\"topic\": ein kurzer slug-artiger Name in Kleinbuchstaben mit Bindestrichen (z.B. \"zulieferer-ch\", \"projekt-alpha\"). NIEMALS eine Zahl.\n" +
        "- entities[].type: GENAU EINER der Werte person, organization, topic — niemals mehrere kombiniert (also NICHT \"person|organization\").\n" +
        "- relations[].type: AUSSCHLIESSLICH einer der oben vorgegebenen Beziehungs-Typen — erfinde KEINE neuen Werte.\n" +
        "- relations[].from und relations[].to MUESSEN exakt einer entities[].id entsprechen.\n" +
        "- Antworte NUR mit dem JSON-Objekt selbst, kein Markdown (kein ```), kein Fliesstext davor oder danach.";

    private const string GraphExtractionExample =
        "BEISPIEL:\n" +
        "Eingabe:\n" +
        "[2024-03-04] MAIL VON anna.berger@zulieferer.ch AN georg@meinefirma.ch: \"Lieferverzug Projekt Alpha\" - Die Teile fuer Projekt Alpha verzoegern sich um zwei Wochen.\n\n" +
        "Korrekte Antwort:\n" +
        "{\n" +
        "  \"entities\": [\n" +
        "    {\"id\":\"anna.berger@zulieferer.ch\",\"type\":\"person\",\"name\":\"Anna Berger\",\"email\":\"anna.berger@zulieferer.ch\"},\n" +
        "    {\"id\":\"georg@meinefirma.ch\",\"type\":\"person\",\"name\":\"Georg\",\"email\":\"georg@meinefirma.ch\"},\n" +
        "    {\"id\":\"zulieferer-ch\",\"type\":\"organization\",\"name\":\"Zulieferer.ch\",\"email\":\"\"},\n" +
        "    {\"id\":\"projekt-alpha\",\"type\":\"topic\",\"name\":\"Projekt Alpha\",\"email\":\"\"}\n" +
        "  ],\n" +
        "  \"relations\": [\n" +
        "    {\"from\":\"anna.berger@zulieferer.ch\",\"to\":\"georg@meinefirma.ch\",\"type\":\"communicated_with\",\"context\":\"Meldet Lieferverzug bei Projekt Alpha\",\"date\":\"2024-03-04\"},\n" +
        "    {\"from\":\"anna.berger@zulieferer.ch\",\"to\":\"zulieferer-ch\",\"type\":\"belongs_to\",\"context\":\"Absender-Domain\",\"date\":\"2024-03-04\"},\n" +
        "    {\"from\":\"anna.berger@zulieferer.ch\",\"to\":\"projekt-alpha\",\"type\":\"works_on\",\"context\":\"Zustaendig fuer Lieferung\",\"date\":\"2024-03-04\"}\n" +
        "  ]\n" +
        "}\n\n";

    // Von GraphExtractionService fuer die serverseitige Normalisierung genutzt (Relation-Typen
    // ausserhalb dieser Listen werden gemappt oder verworfen statt dem Modell blind zu vertrauen).
    public static readonly string[] MailRelationTypes  = ["communicated_with", "works_on", "belongs_to"];
    public static readonly string[] EventRelationTypes = ["attended_meeting", "works_on", "belongs_to"];

    // JSON-Schema fuer Ollamas/OpenAIs structured-output-Modus (response_format.json_schema) —
    // erzwingt Enum-Werte auf Token-Ebene beim Sampling, statt sich nur auf die Prompt-Anweisung
    // zu verlassen (die qwen2.5:14b in der Praxis regelmaessig ignoriert).
    private const string EntitiesSchemaFragment =
        "\"entities\": {\n" +
        "  \"type\": \"array\",\n" +
        "  \"items\": {\n" +
        "    \"type\": \"object\",\n" +
        "    \"properties\": {\n" +
        "      \"id\": { \"type\": \"string\" },\n" +
        "      \"type\": { \"type\": \"string\", \"enum\": [\"person\", \"organization\", \"topic\"] },\n" +
        "      \"name\": { \"type\": \"string\" },\n" +
        "      \"email\": { \"type\": \"string\" }\n" +
        "    },\n" +
        "    \"required\": [\"id\", \"type\", \"name\", \"email\"],\n" +
        "    \"additionalProperties\": false\n" +
        "  }\n" +
        "}";

    private static string RelationsSchemaFragment(IEnumerable<string> relationTypes) =>
        "\"relations\": {\n" +
        "  \"type\": \"array\",\n" +
        "  \"items\": {\n" +
        "    \"type\": \"object\",\n" +
        "    \"properties\": {\n" +
        "      \"from\": { \"type\": \"string\" },\n" +
        "      \"to\": { \"type\": \"string\" },\n" +
        "      \"type\": { \"type\": \"string\", \"enum\": [" + string.Join(",", relationTypes.Select(t => "\"" + t + "\"")) + "] },\n" +
        "      \"context\": { \"type\": \"string\" },\n" +
        "      \"date\": { \"type\": \"string\" }\n" +
        "    },\n" +
        "    \"required\": [\"from\", \"to\", \"type\", \"context\", \"date\"],\n" +
        "    \"additionalProperties\": false\n" +
        "  }\n" +
        "}";

    private static string GraphExtractionSchema(IEnumerable<string> relationTypes) =>
        "{\n" +
        "  \"type\": \"object\",\n" +
        "  \"properties\": {\n" +
        EntitiesSchemaFragment + ",\n" +
        RelationsSchemaFragment(relationTypes) + "\n" +
        "  },\n" +
        "  \"required\": [\"entities\", \"relations\"],\n" +
        "  \"additionalProperties\": false\n" +
        "}";

    public static readonly string MailExtractionSchema  = GraphExtractionSchema(MailRelationTypes);
    public static readonly string EventExtractionSchema = GraphExtractionSchema(EventRelationTypes);

    // "Name <email>" statt nur der Adresse — hilft dem Modell einerseits, sinnvolle entities[].name
    // zu erzeugen, und liefert GraphExtractionService Rohdaten, um erfundene Ids zu korrigieren.
    private static string FormatPerson(string name, string email) => name.Length > 0 ? $"{name} <{email}>" : email;

    // Id-Konvention: bei Personen die E-Mail-Adresse (dient als Merge-Key ueber Batches
    // hinweg), bei organization/topic ein kurzer slug-artiger Name.
    public static string BuildMailExtraction(IEnumerable<GraphMailItem> items)
    {
        var lines = string.Join("\n", items.Select(m =>
            "[" + m.Received.ToString("yyyy-MM-dd") + "] MAIL VON " + FormatPerson(m.FromName, m.From) +
            " AN " + string.Join(",", m.ToRecipients.Select((r, i) => FormatPerson(i < m.ToRecipientNames.Count ? m.ToRecipientNames[i] : "", r))) +
            ": \"" + m.Subject + "\" - " + m.BodyPreview));

        return
            "Du extrahierst Entitaeten und Beziehungen aus geschaeftlicher E-Mail-Korrespondenz fuer einen persoenlichen Knowledge Graph.\n" +
            "Entitaeten-Typen: person, organization, topic. Beziehungs-Typen: communicated_with, works_on, belongs_to.\n\n" +
            GraphExtractionRules + "\n\n" +
            GraphExtractionExample +
            "MAILS:\n" + lines + "\n\n" +
            "Antworte NUR mit validem JSON (kein Markdown), Format:\n" + GraphExtractionFormat;
    }

    public static string BuildEventExtraction(IEnumerable<GraphEventItem> items)
    {
        var lines = string.Join("\n", items.Select(e =>
            "[" + e.Start.ToString("yyyy-MM-dd") + "] MEETING: \"" + e.Subject + "\" mit " +
            string.Join(",", e.Attendees.Select((a, i) => FormatPerson(i < e.AttendeeNames.Count ? e.AttendeeNames[i] : "", a)))));

        return
            "Du extrahierst Entitaeten und Beziehungen aus Kalenderterminen fuer einen persoenlichen Knowledge Graph.\n" +
            "Entitaeten-Typen: person, organization, topic. Beziehungs-Typen: attended_meeting, works_on, belongs_to.\n\n" +
            GraphExtractionRules + "\n\n" +
            GraphExtractionExample +
            "TERMINE:\n" + lines + "\n\n" +
            "Antworte NUR mit validem JSON (kein Markdown), Format:\n" + GraphExtractionFormat;
    }
}
