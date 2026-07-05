namespace tembu_server.Models;

public class GraphMailItem
{
    public string Id { get; set; } = "";
    public string Subject { get; set; } = "";
    public string From { get; set; } = "";
    // Anzeigename des Absenders — Graph liefert ihn im selben "from"-Objekt mit, ohne
    // Zusatzkosten. Dient GraphExtractionService als Nachschlagequelle, um vom LLM erfundene
    // Personen-Ids (statt der vorgegebenen E-Mail-Adresse) zu korrigieren.
    public string FromName { get; set; } = "";
    public List<string> ToRecipients { get; set; } = [];
    // Parallel zu ToRecipients (gleicher Index = gleiche Person), aus demselben Grund wie FromName.
    public List<string> ToRecipientNames { get; set; } = [];
    public DateTimeOffset Received { get; set; }
    public string BodyPreview { get; set; } = "";
    public string ParentFolderId { get; set; } = "";
}

public class GraphEventItem
{
    public string Id { get; set; } = "";
    public string Subject { get; set; } = "";
    public List<string> Attendees { get; set; } = [];
    // Parallel zu Attendees (gleicher Index = gleiche Person) — siehe GraphMailItem.FromName.
    public List<string> AttendeeNames { get; set; } = [];
    public DateTimeOffset Start { get; set; }
    public DateTimeOffset End { get; set; }
}

public class GraphDeltaPage<T>
{
    public List<T> Items { get; set; } = [];
    // Beim nächsten Sync als deltaLink übergeben, um nur Änderungen seit diesem Aufruf
    // zu bekommen statt der ganzen Mailbox/des ganzen Kalenders.
    public string? DeltaLink { get; set; }
}

public class GraphFolderStat
{
    // Wird als excludeFolderIds-Wert an /api/graph/extract zurückgegeben — der User waehlt
    // aus dieser Liste, welche Ordner er aus der Graph-Extraktion ausschliessen will.
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public int TotalItemCount { get; set; }
    public int UnreadItemCount { get; set; }
    // Baumstruktur statt flacher Liste — ein ausgeschlossener Parent schliesst serverseitig
    // (GraphDataService.ExpandWithDescendantsAsync) automatisch alle Children mit ein.
    public List<GraphFolderStat> Children { get; set; } = [];
}

public class GraphMailboxStats
{
    public List<GraphFolderStat> Folders { get; set; } = [];
    public int TotalMailItems { get; set; }
    public int TotalCalendarItems { get; set; }
    // Nur gesetzt, wenn beim Aufruf from/to mitgegeben wurde — reiner Zähl-Query, kein Abruf
    // der eigentlichen Items.
    public int? MailItemsInRange { get; set; }
    public int? CalendarItemsInRange { get; set; }
}
