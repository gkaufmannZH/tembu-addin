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

    public AnalyzeController(AiService ai, UserListService users, GraphAuthService graphAuth)
    {
        _ai        = ai;
        _users     = users;
        _graphAuth = graphAuth;
    }

    // Prüft den Bearer-Token gegen Microsoft Graph statt einem vom Client
    // behaupteten "UserEmail"-Feld zu vertrauen — die E-Mail kommt von Microsoft, nicht vom Client.
    private async Task<(string? email, IActionResult? error)> AuthorizeAsync()
    {
        var header = Request.Headers.Authorization.ToString();
        var token  = header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) ? header["Bearer ".Length..] : null;

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
