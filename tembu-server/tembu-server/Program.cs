using tembu_server.Models;
using tembu_server.Services;

var builder = WebApplication.CreateBuilder(args);

// Separate, vom Admin lokal editierbare Datei — Live-Reload, damit Änderungen ohne
// Server-Neustart wirken (siehe AiService, der IOptionsMonitor<AiSettings> nutzt).
builder.Configuration.AddJsonFile("ai-settings.json", optional: true, reloadOnChange: true);
builder.Services.Configure<AiSettings>(builder.Configuration.GetSection("AI"));

// Lokale, nicht eingecheckte Overrides (echte Lizenzwerte) — appsettings.json
// enthält nur Platzhalter, damit kein echter Key ins Repo gelangt.
builder.Configuration.AddJsonFile("appsettings.Local.json", optional: true, reloadOnChange: true);

builder.Services.AddControllers();
builder.Services.AddSingleton<AiService>();
builder.Services.AddSingleton<UserListService>();
builder.Host.UseWindowsService();

builder.Services.AddCors(o => o.AddDefaultPolicy(p => p
    .WithOrigins(
        "https://localhost:5001",
        "https://outlook.office.com",
        "https://outlook.office365.com",
        "https://outlook.live.com",
        "https://gkaufmannzh.github.io")
    .AllowAnyHeader()
    .AllowAnyMethod()));

// ── Lizenzprüfung beim Start ───────────────────────────────────────────────
var cfg      = builder.Configuration;
var email  = cfg["License:Email"]  ?? "";
var key    = cfg["License:Key"]    ?? "";
var expiry = cfg["License:Expiry"] ?? "";

var license = LicenseService.Validate(email, key, expiry);
if (!license.IsValid)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"TEMBU LIZENZFEHLER: {license.ErrorMessage}");
    Console.ResetColor();
    Environment.Exit(1);
}

Console.WriteLine($"Tembu Server — Lizenz gültig bis {license.ExpiryDate:dd.MM.yyyy} ({license.DaysRemaining} Tage)");

// LicenseInfo als Singleton bereitstellen (für StatusController)
builder.Services.AddSingleton(license);

// ── App aufbauen ───────────────────────────────────────────────────────────
var app = builder.Build();

app.UseHttpsRedirection();
app.UseCors();
app.UseDefaultFiles();
app.UseStaticFiles();
app.UseAuthorization();
app.MapControllers();

app.Run();
