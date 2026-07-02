using Microsoft.AspNetCore.Mvc;
using tembu_server.Models;
using tembu_server.Services;

namespace tembu_server.Controllers;

[ApiController]
[Route("api/analyze")]
public class AnalyzeController : ControllerBase
{
    private readonly AiService _ai;
    private readonly UserListService _users;
    private readonly GraphAuthService _graphAuth;
    private readonly ILogger<AnalyzeController> _logger;

    public AnalyzeController(AiService ai, UserListService users, GraphAuthService graphAuth, ILogger<AnalyzeController> logger)
    {
        _ai        = ai;
        _users     = users;
        _graphAuth = graphAuth;
        _logger    = logger;
    }

    // Prüft den Bearer-Token gegen Microsoft Graph statt einem vom Client
    // behaupteten "UserEmail"-Feld zu vertrauen — die E-Mail kommt von Microsoft, nicht vom Client.
    private async Task<(string? email, IActionResult? error)> AuthorizeAsync()
    {
        var header = Request.Headers.Authorization.ToString();
        var token  = header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? header["Bearer ".Length..] : null;

        if (string.IsNullOrEmpty(token))
            _logger.LogWarning("AnalyzeController: kein/leerer Authorization-Header. Vorhandene Header: {Headers} | Origin: {Origin}",
                string.Join(", ", Request.Headers.Keys), Request.Headers.Origin.ToString());

        var email = await _graphAuth.GetVerifiedEmailAsync(token);
        if (email == null)
            return (null, Unauthorized("Ungültiger oder fehlender Microsoft-Token."));

        if (!_users.IsAuthorized(email))
            return (null, StatusCode(403, $"Nutzer '{email}' ist nicht lizenziert."));

        return (email, null);
    }

    [HttpPost]
    public async Task<IActionResult> Analyze([FromBody] ContactData data)
    {
        var (_, error) = await AuthorizeAsync();
        if (error != null) return error;

        try
        {
            var prompt = PromptBuilder.Build(data);
            var result = await _ai.CallAsync(prompt);
            return Ok(result);
        }
        catch (Exception ex)
        {
            return StatusCode(500, ex.Message);
        }
    }

    [HttpPost("company")]
    public async Task<IActionResult> AnalyzeCompany([FromBody] CompanyData data)
    {
        var (_, error) = await AuthorizeAsync();
        if (error != null) return error;

        try
        {
            var prompt = PromptBuilder.BuildCompany(data);
            var result = await _ai.CallAsync(prompt);
            return Ok(result);
        }
        catch (Exception ex)
        {
            return StatusCode(500, ex.Message);
        }
    }
}
