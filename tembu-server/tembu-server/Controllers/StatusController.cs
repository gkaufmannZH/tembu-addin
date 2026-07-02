using Microsoft.AspNetCore.Mvc;
using tembu_server.Models;
using tembu_server.Services;

namespace tembu_server.Controllers;

[ApiController]
[Route("api/status")]
public class StatusController : ControllerBase
{
    private readonly LicenseInfo _license;
    private readonly UserListService _users;

    public StatusController(LicenseInfo license, UserListService users)
    {
        _license = license;
        _users   = users;
    }

    // Unauthentifiziert erreichbar (Health-Check) — daher keine identifizierenden
    // Lizenzdaten (E-Mail) preisgeben, nur Betriebszustand.
    [HttpGet]
    public IActionResult Get() => Ok(new
    {
        status           = "ok",
        version          = "2.0",
        licenseExpiry    = _license.ExpiryDate.ToString("yyyy-MM-dd"),
        daysRemaining    = _license.DaysRemaining,
        authorizedUsers  = _users.UserCount,
    });
}
