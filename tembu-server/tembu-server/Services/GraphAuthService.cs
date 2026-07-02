using System.Net.Http.Headers;
using System.Text.Json;

namespace tembu_server.Services;

// Prüft den vom Client mitgeschickten Microsoft-Graph-Token direkt bei Microsoft
// und liefert die verifizierte E-Mail zurück — der Client kann sich NICHT mehr
// per beliebig gewähltem "UserEmail"-Feld als jemand anderes ausgeben.
public class GraphAuthService
{
    private readonly HttpClient _http = new();

    public async Task<string?> GetVerifiedEmailAsync(string? bearerToken)
    {
        if (string.IsNullOrWhiteSpace(bearerToken)) return null;

        using var req = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

        try
        {
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;

            var json = await res.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var mail = doc.RootElement.TryGetProperty("mail", out var m) ? m.GetString() : null;
            var upn  = doc.RootElement.TryGetProperty("userPrincipalName", out var u) ? u.GetString() : null;
            var email = mail ?? upn;
            return string.IsNullOrWhiteSpace(email) ? null : email.ToLowerInvariant().Trim();
        }
        catch
        {
            return null;
        }
    }
}
