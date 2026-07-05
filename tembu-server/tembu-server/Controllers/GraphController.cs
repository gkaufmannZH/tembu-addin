using Microsoft.AspNetCore.Mvc;
using tembu_server.Models;
using tembu_server.Services;

namespace tembu_server.Controllers;

// Testendpoint fuer den Knowledge-Graph-Aufbau: zieht Mail-/Kalenderdaten des eingeloggten
// Users per Graph-Delta-Query und laesst sie vom konfigurierten LLM zu Entitaeten/Beziehungen
// extrahieren. Noch ohne Persistenz — jeder Aufruf ohne deltaLink-Query-Parameter verarbeitet
// den aktuellen Stand neu, der zurueckgegebene deltaLink kann fuer einen inkrementellen
// Folgeaufruf mitgegeben werden.
[ApiController]
[Route("api/graph")]
public class GraphController(GraphExtractionService extraction, GraphDataService graphData, GraphAuthService graphAuth, UserListService users, ILogger<GraphController> logger) : ControllerBase
{
    // Gleiches Muster wie AnalyzeController.AuthorizeAsync, gibt zusaetzlich den rohen Token
    // zurueck — GraphExtractionService braucht ihn fuer eigene Graph-API-Aufrufe, nicht nur
    // fuer die einmalige Identitaetspruefung.
    private async Task<(string? email, string? token, IActionResult? error)> AuthorizeAsync()
    {
        var header = Request.Headers.Authorization.ToString();
        var token  = header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? header["Bearer ".Length..] : null;

        var email = await graphAuth.GetVerifiedEmailAsync(token);
        if (email == null)
            return (null, null, Unauthorized("Ungültiger oder fehlender Microsoft-Token."));

        if (!users.IsAuthorized(email))
            return (null, null, StatusCode(403, $"Nutzer '{email}' ist nicht lizenziert."));

        return (email, token, null);
    }

    // Vor dem ersten vollen Extract-Aufruf: grobe Größenschätzung (Item-Anzahl pro Mailordner
    // + Terminanzahl), damit man abschätzen kann, wie lange/teuer ein Erstlauf wird.
    // Optional from/to (ISO 8601) liefert zusätzlich eine reine Zählung (kein Datenabruf) fuer
    // genau dieses Zeitfenster.
    [HttpGet("stats")]
    public async Task<IActionResult> Stats([FromQuery] DateTimeOffset? from, [FromQuery] DateTimeOffset? to)
    {
        var (email, token, error) = await AuthorizeAsync();
        if (error != null) return error;

        try
        {
            var stats = await graphData.GetMailboxStatsAsync(token!, from, to);
            logger.LogInformation("GraphStats: {Email} — {MailItems} Mails, {EventItems} Termine",
                email, stats.TotalMailItems, stats.TotalCalendarItems);
            return Ok(stats);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "GraphStats fehlgeschlagen fuer {Email}", email);
            return StatusCode(500, ex.Message);
        }
    }

    // Als POST-Body statt Query-Parameter (siehe GraphExtractRequest) — ExcludeFolderIds kann
    // durch Parent-Ordner-Kaskadierung leicht mehrere hundert IDs umfassen, das sprengt jede
    // Request-URL-Laenge. from/to greifen nur beim Erstlauf (kein deltaLink); ExcludeFolderIds
    // weggelassen = Default (Papierkorb + Junk automatisch ausgeschlossen).
    [HttpPost("extract")]
    public async Task<IActionResult> Extract([FromBody] GraphExtractRequest? request)
    {
        var (email, token, error) = await AuthorizeAsync();
        if (error != null) return error;

        var req = request ?? new GraphExtractRequest();

        try
        {
            var result = await extraction.ExtractAsync(token!, req.MailDeltaLink, req.CalendarDeltaLink, req.From, req.To, req.ExcludeFolderIds);
            logger.LogInformation("GraphExtraction: {Email} — {EntityCount} Entitaeten, {RelationCount} Beziehungen",
                email, result.Entities.Count, result.Relations.Count);
            return Ok(result);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "GraphExtraction fehlgeschlagen fuer {Email}", email);
            return StatusCode(500, ex.Message);
        }
    }
}
