using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.Extensions.Options;
using tembu_server.Models;

namespace tembu_server.Services;

// Baut die CORS-Policy bei jedem Request neu aus IOptionsMonitor<CorsSettings> —
// dadurch wirkt eine Änderung an cors-settings.json sofort, ohne Server-Neustart.
public class DynamicCorsPolicyProvider(IOptionsMonitor<CorsSettings> settings) : ICorsPolicyProvider
{
    public Task<CorsPolicy?> GetPolicyAsync(HttpContext context, string? policyName)
    {
        var origins = settings.CurrentValue.AllowedOrigins;
        var policy = new CorsPolicyBuilder(origins)
            .AllowAnyHeader()
            .AllowAnyMethod()
            .Build();
        return Task.FromResult<CorsPolicy?>(policy);
    }
}
