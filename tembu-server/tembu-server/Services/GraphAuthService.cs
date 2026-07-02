using System.Net.Http.Headers;
using System.Text.Json;

namespace tembu_server.Services;

// Prüft den vom Client mitgeschickten Microsoft-Graph-Token direkt bei Microsoft
// und liefert die verifizierte E-Mail zurück — der Client kann sich NICHT mehr
// per beliebig gewähltem "UserEmail"-Feld als jemand anderes ausgeben.
public class GraphAuthService(ILogger<GraphAuthService> logger)
{
    private readonly HttpClient _http = new();

    public async Task<string?> GetVerifiedEmailAsync(string? bearerToken)
    {
        if (string.IsNullOrWhiteSpace(bearerToken))
        {
            logger.LogWarning("GraphAuth: kein Bearer-Token im Request");
            return null;
        }

        using var req = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

        try
        {
            var res = await _http.SendAsync(req);
            var json = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
            {
                logger.LogWarning("GraphAuth: /me lieferte {Status} — {Body}", (int)res.StatusCode, json.Length > 300 ? json[..300] : json);
                return null;
            }

            using var doc = JsonDocument.Parse(json);
            var mail = doc.RootElement.TryGetProperty("mail", out var m) ? m.GetString() : null;
            var upn  = doc.RootElement.TryGetProperty("userPrincipalName", out var u) ? u.GetString() : null;
            var email = mail ?? upn;
            if (string.IsNullOrWhiteSpace(email))
                logger.LogWarning("GraphAuth: /me OK, aber weder mail noch userPrincipalName gesetzt");
            return string.IsNullOrWhiteSpace(email) ? null : email.ToLowerInvariant().Trim();
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "GraphAuth: Aufruf von /me fehlgeschlagen");
            return null;
        }
    }
}
