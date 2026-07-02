using tembu_server.Models;

namespace tembu_server.Services;

public static class PromptBuilder
{
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
            "  \"sentiment\": \"positiv|neutral|negativ\",\n" +
            "  \"openPoints\": [\"Offener Punkt 1\"],\n" +
            "  \"themes\": [\n" +
            "    { \"name\": \"Thema\", \"status\": \"offen|abgeschlossen\", \"summary\": \"Kurzbeschreibung\",\n" +
            "      \"interactions\": [{\"date\":\"YYYY-MM-DD\",\"type\":\"email|meeting\",\"subject\":\"Betreff\"}] }\n" +
            "  ],\n" +
            "  \"nextStep\": \"Konkrete Empfehlung fuer naechstes Gespraech\",\n" +
            "  \"background\": \"Oeffentlich bekannte Infos zu " + data.ContactName + ". Falls unbekannt: leer lassen.\"\n" +
            "}";

        return
            "Du bist ein persoenlicher Business-Assistent. Heute ist " + today + ".\n" +
            "Analysiere alle Interaktionen mit \"" + data.ContactName + "\" (" + data.ContactEmail + ").\n\n" +
            blocks +
            "Antworte NUR mit validem JSON (kein Markdown).\n" +
            "Wichtig: Im interactions-Array jedes Themas ALLE zugehoerigen Interaktionen auflisten.\n" +
            jsonTemplate;
    }
}
