using System.Net.Http.Headers;
using System.Text.Json;
using tembu_server.Models;

namespace tembu_server.Services;

// Zieht Mail- und Kalenderdaten per Microsoft-Graph-Delta-Query für den Knowledge-Graph-Aufbau.
// Läuft mit dem Bearer-Token des Users (delegated permissions, derselbe Token wie in
// GraphAuthService) — der Server braucht keine eigenen Graph-Credentials und sieht nur,
// wozu der eingeloggte User selbst Zugriff hat.
//
// Delta-Query-Prinzip: der erste Aufruf ohne deltaLink liefert (paginiert) alles im
// aktuellen Stand + am Ende einen deltaLink. Jeder folgende Aufruf mit diesem Link liefert
// nur noch Änderungen seit dem letzten Sync — kein Reprocessing der ganzen Mailbox/des
// ganzen Kalenders bei jedem Lauf.
public class GraphDataService(ILogger<GraphDataService> logger)
{
    private const string Base = "https://graph.microsoft.com/v1.0";

    // Default-Timeout (100s) reicht bei grossen Zeitfenstern/Mailboxen mit vielen Seiten nicht
    // immer — einzelne Graph-Seitenabrufe koennen bei Throttling laenger dauern.
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromMinutes(5) };

    // Graph unterstuetzt KEIN Delta auf /me/messages ohne Ordner-Scope ("Change tracking is
    // not supported against 'microsoft.graph.message'", 400 BadRequest) — Delta-Query fuer
    // Mails geht nur pro Mailordner. Fuer mailbox-weiten Zugriff bleibt deshalb nur die normale
    // (nicht-inkrementelle) Listing-Query mit $filter auf from/to; deltaLink wird fuer Mails
    // daher nie gesetzt (GraphDeltaPage.DeltaLink bleibt null) — jeder Aufruf fragt das
    // Zeitfenster komplett neu ab statt nur Aenderungen seit dem letzten Sync. Fuer echtes
    // inkrementelles Mail-Sync muesste man pro Ordner einzeln per delta pollen.
    // excludeFolderIds: vom User gewaehlte Ordner-IDs (aus GraphMailboxStats.Folders), die
    // NICHT in den Graphen einfliessen sollen. null = Default (Papierkorb + Junk automatisch
    // ausschliessen), leere Liste = User hat sich bewusst gegen jeden Ausschluss entschieden.
    public async Task<GraphDeltaPage<GraphMailItem>> GetMailDeltaAsync(string bearerToken, string? deltaLink = null, DateTimeOffset? from = null, DateTimeOffset? to = null, IReadOnlyCollection<string>? excludeFolderIds = null)
    {
        string startUrl;
        if (deltaLink != null)
        {
            startUrl = deltaLink;
        }
        else
        {
            // $select begrenzt bewusst auf das für die Extraktion Nötige (kein voller Body) —
            // Datenminimierung, nicht nur Payload-Größe.
            var filters = new List<string>();
            if (from.HasValue) filters.Add($"receivedDateTime ge {ToGraphDateTime(from.Value)}");
            if (to.HasValue)   filters.Add($"receivedDateTime le {ToGraphDateTime(to.Value)}");
            var filterPart = filters.Count > 0 ? $"&$filter={Uri.EscapeDataString(string.Join(" and ", filters))}" : "";

            startUrl = $"{Base}/me/messages?$select=subject,from,toRecipients,receivedDateTime,bodyPreview,parentFolderId&$top=50&$orderby=receivedDateTime desc{filterPart}";
        }

        var page = await FetchAllPagesAsync(bearerToken, startUrl, ParseMailItem);

        // Delta-Query kennt keinen Ordner-Ausschlussfilter — deshalb clientseitig nach dem
        // Fetch rausfiltern statt serverseitig einzuschränken.
        var excluded = excludeFolderIds != null
            ? await ExpandWithDescendantsAsync(bearerToken, excludeFolderIds)
            : await GetDefaultExcludedFolderIdsAsync(bearerToken);

        if (excluded.Count > 0)
        {
            var before = page.Items.Count;
            page.Items.RemoveAll(m => excluded.Contains(m.ParentFolderId));
            logger.LogInformation("GraphData: {Removed} Mails aus ausgeschlossenen Ordnern gefiltert (von {Before})", before - page.Items.Count, before);
        }

        return page;
    }

    // Schliesst ein Parent-Ordner den User aus, sollen automatisch auch alle Unterordner
    // ausgeschlossen sein. Holt dafuer EINMAL den ganzen Ordnerbaum (wie GetMailFolderTreeAsync
    // fuer /stats) und expandiert lokal im Speicher — nicht pro auszuschliessender ID einen
    // eigenen Graph-Call (das waren bei vielen ausgewaehlten Ordnern hunderte sequenzielle
    // Requests und hat Graph-Throttling ausgeloest, das wiederum einen einzelnen Request
    // ueber das 100s-HttpClient-Timeout laufen liess).
    private async Task<HashSet<string>> ExpandWithDescendantsAsync(string bearerToken, IEnumerable<string> folderIds)
    {
        var tree = await GetMailFolderTreeAsync(bearerToken);
        var childrenOf = new Dictionary<string, List<string>>();
        CollectChildrenMap(tree, childrenOf);

        var result = new HashSet<string>(folderIds);
        var queue = new Queue<string>(result);

        while (queue.Count > 0)
        {
            var id = queue.Dequeue();
            if (!childrenOf.TryGetValue(id, out var children)) continue;
            foreach (var childId in children)
                if (result.Add(childId)) queue.Enqueue(childId);
        }

        return result;
    }

    private static void CollectChildrenMap(List<GraphFolderStat> folders, Dictionary<string, List<string>> map)
    {
        foreach (var f in folders)
        {
            map[f.Id] = f.Children.Select(c => c.Id).ToList();
            CollectChildrenMap(f.Children, map);
        }
    }

    // Fallback, solange der User noch keine eigene Ordnerauswahl getroffen hat (siehe
    // GraphController.Stats fuer die Auswahlliste selbst).
    private async Task<HashSet<string>> GetDefaultExcludedFolderIdsAsync(string bearerToken)
    {
        var ids = new HashSet<string>();
        foreach (var wellKnownName in new[] { "deleteditems", "junkemail" })
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, $"{Base}/me/mailFolders/{wellKnownName}?$select=id");
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) continue; // z.B. kein separater Junk-Ordner in dieser Mailbox

            using var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            if (doc.RootElement.TryGetProperty("id", out var idEl) && idEl.GetString() is { } id)
                ids.Add(id);
        }
        return ids;
    }

    public Task<GraphDeltaPage<GraphEventItem>> GetCalendarDeltaAsync(string bearerToken, string? deltaLink = null, DateTimeOffset? from = null, DateTimeOffset? to = null)
    {
        string startUrl;
        if (deltaLink != null)
        {
            startUrl = deltaLink;
        }
        else
        {
            // calendarView/delta statt events/delta: braucht zwingend ein Zeitfenster, expandiert
            // dafür wiederkehrende Termine korrekt zu Einzelinstanzen innerhalb des Fensters —
            // reines events/delta kann das nicht und eignet sich nicht für einen "von-bis"-Filter.
            var f = from ?? DateTimeOffset.UtcNow.AddMonths(-6);
            var t = to   ?? DateTimeOffset.UtcNow;
            startUrl = $"{Base}/me/calendarView/delta?startDateTime={ToGraphDateTime(f)}&endDateTime={ToGraphDateTime(t)}&$select=subject,attendees,start,end";
        }

        return FetchAllPagesAsync(bearerToken, startUrl, ParseEventItem);
    }

    private static string ToGraphDateTime(DateTimeOffset dt) => dt.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");

    // Teams-Chatnachrichten brauchen den Scope "Chat.Read", den outlook-addin/auth.html aktuell
    // nicht anfragt — erst nach Scope-Erweiterung + erneutem User-Consent nutzbar
    // (GET /me/chats/{id}/messages/delta, kein globaler Delta-Endpoint über alle Chats).

    // Grobe Größenschätzung vor dem ersten vollen Extraktions-Lauf — reine Mail-/Termin-Anzahl,
    // damit man abschätzen kann, wie viele LLM-Batches ein Erstlauf erzeugt, ohne die Mailbox
    // schon zu verarbeiten.
    public async Task<GraphMailboxStats> GetMailboxStatsAsync(string bearerToken, DateTimeOffset? from = null, DateTimeOffset? to = null)
    {
        var folders = await GetMailFolderTreeAsync(bearerToken);
        var stats = new GraphMailboxStats
        {
            Folders             = folders,
            TotalMailItems      = SumTotalItems(folders),
            TotalCalendarItems  = await GetCalendarCountAsync(bearerToken),
        };

        if (from.HasValue || to.HasValue)
        {
            stats.MailItemsInRange     = await GetMailCountInRangeAsync(bearerToken, from, to);
            stats.CalendarItemsInRange = await GetCalendarCountInRangeAsync(bearerToken, from, to);
        }

        return stats;
    }

    // Reiner $count-Query über /me/messages (ordnerübergreifend) mit demselben receivedDateTime-
    // Filter wie GetMailDeltaAsync — zaehlt, ohne die Items selbst abzurufen.
    private async Task<int> GetMailCountInRangeAsync(string bearerToken, DateTimeOffset? from, DateTimeOffset? to)
    {
        var filters = new List<string>();
        if (from.HasValue) filters.Add($"receivedDateTime ge {ToGraphDateTime(from.Value)}");
        if (to.HasValue)   filters.Add($"receivedDateTime le {ToGraphDateTime(to.Value)}");
        var filterPart = filters.Count > 0 ? $"&$filter={Uri.EscapeDataString(string.Join(" and ", filters))}" : "";

        using var req = new HttpRequestMessage(HttpMethod.Get, $"{Base}/me/messages?$top=1&$count=true{filterPart}");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
        req.Headers.Add("ConsistencyLevel", "eventual");

        var res  = await _http.SendAsync(req);
        var json = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            logger.LogWarning("GraphData: Mail-Count(Zeitraum) fehlgeschlagen {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
            return 0;
        }

        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("@odata.count", out var c) ? c.GetInt32() : 0;
    }

    private async Task<int> GetCalendarCountInRangeAsync(string bearerToken, DateTimeOffset? from, DateTimeOffset? to)
    {
        var f = from ?? DateTimeOffset.UtcNow.AddMonths(-6);
        var t = to   ?? DateTimeOffset.UtcNow;

        using var req = new HttpRequestMessage(HttpMethod.Get, $"{Base}/me/calendarView?startDateTime={ToGraphDateTime(f)}&endDateTime={ToGraphDateTime(t)}&$top=1&$count=true");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
        req.Headers.Add("ConsistencyLevel", "eventual");

        var res  = await _http.SendAsync(req);
        var json = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            logger.LogWarning("GraphData: Termin-Count(Zeitraum) fehlgeschlagen {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
            return 0;
        }

        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("@odata.count", out var c) ? c.GetInt32() : 0;
    }

    // Baut den Ordnerbaum rekursiv auf (Top-Level + alle Unterordner), damit die UI eine
    // Baumansicht zeigen kann und ein ausgeschlossener Parent seine Children mit ausschliesst.
    private Task<List<GraphFolderStat>> GetMailFolderTreeAsync(string bearerToken) =>
        GetFolderLevelAsync(bearerToken, $"{Base}/me/mailFolders?$top=250&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount");

    private async Task<List<GraphFolderStat>> GetFolderLevelAsync(string bearerToken, string url)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

        var res  = await _http.SendAsync(req);
        var json = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            logger.LogWarning("GraphData: Mailbox-Stats fehlgeschlagen {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
            throw new HttpRequestException($"Graph-Abruf fehlgeschlagen ({(int)res.StatusCode})");
        }

        using var doc = JsonDocument.Parse(json);
        var folders = new List<GraphFolderStat>();
        foreach (var f in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var id = GetString(f, "id");
            var node = new GraphFolderStat
            {
                Id              = id,
                Name            = GetString(f, "displayName"),
                TotalItemCount  = f.TryGetProperty("totalItemCount", out var t) ? t.GetInt32() : 0,
                UnreadItemCount = f.TryGetProperty("unreadItemCount", out var u) ? u.GetInt32() : 0,
            };

            var childCount = f.TryGetProperty("childFolderCount", out var cc) ? cc.GetInt32() : 0;
            if (childCount > 0)
                node.Children = await GetFolderLevelAsync(bearerToken, $"{Base}/me/mailFolders/{id}/childFolders?$top=250&$select=id,displayName,totalItemCount,unreadItemCount,childFolderCount");

            folders.Add(node);
        }
        return folders;
    }

    private static int SumTotalItems(IEnumerable<GraphFolderStat> folders) =>
        folders.Sum(f => f.TotalItemCount + SumTotalItems(f.Children));

    private async Task<int> GetCalendarCountAsync(string bearerToken)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, $"{Base}/me/events?$top=1&$count=true");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
        req.Headers.Add("ConsistencyLevel", "eventual"); // von Graph für $count auf /events verlangt

        var res  = await _http.SendAsync(req);
        var json = await res.Content.ReadAsStringAsync();
        if (!res.IsSuccessStatusCode)
        {
            // Termin-Anzahl ist nur eine Zusatzschätzung — nicht die Mail-Stats deswegen verwerfen.
            logger.LogWarning("GraphData: Kalender-Count fehlgeschlagen {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
            return 0;
        }

        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("@odata.count", out var c) ? c.GetInt32() : 0;
    }

    private async Task<GraphDeltaPage<T>> FetchAllPagesAsync<T>(string bearerToken, string startUrl, Func<JsonElement, T> parse)
    {
        var items = new List<T>();
        string? nextUrl = startUrl;
        string? deltaLink = null;

        while (nextUrl != null)
        {
            var (res, json) = await SendWithRetryAsync(bearerToken, nextUrl);

            if (!res.IsSuccessStatusCode)
            {
                logger.LogWarning("GraphData: Abruf fehlgeschlagen {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
                throw new HttpRequestException($"Graph-Abruf fehlgeschlagen ({(int)res.StatusCode})");
            }

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            foreach (var el in root.GetProperty("value").EnumerateArray())
            {
                if (el.TryGetProperty("@removed", out _)) continue; // gelöscht/verschoben seit letztem Sync — kein Inhalt mehr abrufbar
                items.Add(parse(el));
            }

            nextUrl = root.TryGetProperty("@odata.nextLink", out var next) ? next.GetString() : null;
            if (nextUrl == null && root.TryGetProperty("@odata.deltaLink", out var delta))
                deltaLink = delta.GetString();
        }

        return new GraphDeltaPage<T> { Items = items, DeltaLink = deltaLink };
    }

    // Graph reisst bei langen paginierten Abrufen (viele Seiten, z.B. ein Jahr Mailbox-weiter
    // Mails) gelegentlich die Verbindung ab (SocketException "Remotehost geschlossen") —
    // transiente Netzwerkfehler hier mit kurzem Backoff erneut versuchen statt den ganzen
    // Extraktionslauf abzubrechen.
    private async Task<(HttpResponseMessage res, string json)> SendWithRetryAsync(string bearerToken, string url, int maxAttempts = 3)
    {
        for (var attempt = 1; ; attempt++)
        {
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

                var res = await _http.SendAsync(req);
                var json = await res.Content.ReadAsStringAsync();
                return (res, json);
            }
            catch (HttpRequestException ex) when (attempt < maxAttempts)
            {
                logger.LogWarning(ex, "GraphData: transienter Fehler bei Versuch {Attempt}/{Max}, retry nach Backoff", attempt, maxAttempts);
                await Task.Delay(TimeSpan.FromSeconds(attempt * 2));
            }
        }
    }

    private static GraphMailItem ParseMailItem(JsonElement m)
    {
        (string address, string name) from = m.TryGetProperty("from", out var f) ? ParseEmailAddress(f) : ("", "");
        var to = m.TryGetProperty("toRecipients", out var tr) ? tr.EnumerateArray().Select(ParseEmailAddress).ToList() : [];

        return new GraphMailItem
        {
            Id = m.GetProperty("id").GetString() ?? "",
            Subject = GetString(m, "subject"),
            From = from.address,
            FromName = from.name,
            ToRecipients = to.Select(x => x.address).ToList(),
            ToRecipientNames = to.Select(x => x.name).ToList(),
            Received = m.TryGetProperty("receivedDateTime", out var rd) && rd.ValueKind == JsonValueKind.String ? rd.GetDateTimeOffset() : default,
            BodyPreview = GetString(m, "bodyPreview"),
            ParentFolderId = GetString(m, "parentFolderId"),
        };
    }

    private static GraphEventItem ParseEventItem(JsonElement e)
    {
        var attendees = e.TryGetProperty("attendees", out var att) ? att.EnumerateArray().Select(ParseEmailAddress).ToList() : [];

        return new GraphEventItem
        {
            Id = e.GetProperty("id").GetString() ?? "",
            Subject = GetString(e, "subject"),
            Attendees = attendees.Select(a => a.address).ToList(),
            AttendeeNames = attendees.Select(a => a.name).ToList(),
            Start = e.TryGetProperty("start", out var st) ? ParseGraphDateTime(st) : default,
            End = e.TryGetProperty("end", out var en) ? ParseGraphDateTime(en) : default,
        };
    }

    // "from", je ein toRecipients[]- und attendees[]-Eintrag haben alle dieselbe Form
    // { "emailAddress": { "address": "...", "name": "..." } } — eine Parse-Funktion fuer alle drei.
    private static (string address, string name) ParseEmailAddress(JsonElement el) =>
        el.TryGetProperty("emailAddress", out var ea) ? (GetString(ea, "address"), GetString(ea, "name")) : ("", "");

    private static DateTimeOffset ParseGraphDateTime(JsonElement dt) =>
        DateTimeOffset.TryParse(GetString(dt, "dateTime"), out var parsed) ? parsed : default;

    private static string GetString(JsonElement el, string prop) =>
        el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
}
