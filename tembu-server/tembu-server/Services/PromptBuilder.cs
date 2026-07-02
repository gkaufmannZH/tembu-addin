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
}
