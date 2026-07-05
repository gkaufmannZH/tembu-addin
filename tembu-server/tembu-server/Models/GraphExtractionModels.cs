namespace tembu_server.Models;

public class GraphEntity
{
    // Bei Personen die E-Mail-Adresse, sonst ein kurzer slug-artiger Name — dient als Merge-Key
    // ueber mehrere Extraktions-Batches hinweg (z.B. dieselbe Person aus Mail- und Kalenderdaten).
    public string Id { get; set; } = "";
    public string Type { get; set; } = ""; // person | organization | topic
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
}

public class GraphRelation
{
    public string From { get; set; } = "";
    public string To { get; set; } = "";
    public string Type { get; set; } = ""; // communicated_with | attended_meeting | works_on | belongs_to
    public string Context { get; set; } = "";
    public string Date { get; set; } = "";
}

public class GraphExtractionResult
{
    public List<GraphEntity> Entities { get; set; } = [];
    public List<GraphRelation> Relations { get; set; } = [];
    public string? MailDeltaLink { get; set; }
    public string? CalendarDeltaLink { get; set; }
}

// Als POST-Body statt Query-Parameter, weil ExcludeFolderIds bei vielen ausgeschlossenen
// Ordnern (inkl. kaskadierter Unterordner) leicht auf mehrere hundert IDs anwaechst — das
// sprengt Request-URL-Laengenlimits (Browser, Server, Kommandozeile).
public class GraphExtractRequest
{
    public string? MailDeltaLink { get; set; }
    public string? CalendarDeltaLink { get; set; }
    public DateTimeOffset? From { get; set; }
    public DateTimeOffset? To { get; set; }
    public List<string>? ExcludeFolderIds { get; set; }
}
