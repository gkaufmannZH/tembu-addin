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

    public AnalyzeController(AiService ai, UserListService users)
    {
        _ai    = ai;
        _users = users;
    }

    [HttpPost]
    public async Task<IActionResult> Analyze([FromBody] ContactData data)
    {
        if (!_users.IsAuthorized(data.UserEmail))
            return StatusCode(403, $"Nutzer '{data.UserEmail}' ist nicht lizenziert.");

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
}
