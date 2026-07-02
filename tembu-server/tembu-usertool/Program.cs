using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

// ── Konstanten (identisch mit tembu-server) ──────────────────────────────────
const string Secret = "tembu-2024-x7k9-secure";
const string Prefix = "TEMBU";

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.Title = "Tembu User Tool";
Console.ForegroundColor = ConsoleColor.Cyan;
Console.WriteLine("╔══════════════════════════════════════╗");
Console.WriteLine("║     TEMBU USER TOOL  v2.1            ║");
Console.WriteLine("╚══════════════════════════════════════╝");
Console.ResetColor();
Console.WriteLine();

// ── Modus wählen ─────────────────────────────────────────────────────────────
Console.WriteLine("Was möchten Sie tun?");
Console.WriteLine("  [1]  Neue users.dat erstellen");
Console.WriteLine("  [2]  Lizenzschlüssel generieren");
Console.WriteLine("  [3]  Bestehende users.dat anzeigen");
Console.WriteLine("  [4]  KI-Einstellungen bearbeiten (ai-settings.json)");
Console.Write("\nAuswahl: ");
var mode = Console.ReadLine()?.Trim();
Console.WriteLine();

switch (mode)
{
    case "1": CreateUserFile(); break;
    case "2": GenerateLicenseKey(); break;
    case "3": ShowUserFile(); break;
    case "4": EditAiSettings(); break;
    default:
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("Ungültige Auswahl.");
        Console.ResetColor();
        break;
}

Console.WriteLine();
Console.WriteLine("Drücken Sie eine Taste zum Beenden...");
Console.ReadKey();

// ── Funktionen ────────────────────────────────────────────────────────────────

void CreateUserFile()
{
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine("─── Neue users.dat erstellen ───────────────────────────");
    Console.ResetColor();

    Console.Write("Kundenname (z.B. Firma AG): ");
    var customer = Console.ReadLine()?.Trim() ?? "";

    Console.WriteLine();
    Console.WriteLine("Geben Sie die E-Mail-Adressen ein (eine pro Zeile, leer lassen zum Beenden):");
    var emails = new List<string>();
    while (true)
    {
        Console.Write($"  [{emails.Count + 1}] ");
        var email = Console.ReadLine()?.Trim();
        if (string.IsNullOrEmpty(email)) break;
        if (!email.Contains('@'))
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("       → Ungültige E-Mail, übersprungen.");
            Console.ResetColor();
            continue;
        }
        emails.Add(email.ToLower());
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"       → OK");
        Console.ResetColor();
    }

    if (emails.Count == 0)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("Keine E-Mails eingegeben. Abbruch.");
        Console.ResetColor();
        return;
    }

    Console.Write("\nAusgabepfad (Enter = users.dat im aktuellen Verzeichnis): ");
    var path = Console.ReadLine()?.Trim();
    if (string.IsNullOrEmpty(path)) path = "users.dat";

    var userList = new UserListData
    {
        Customer = customer,
        Created  = DateTime.Today.ToString("yyyy-MM-dd"),
        Users    = emails,
    };

    var encrypted = Encrypt(userList);
    File.WriteAllBytes(path, encrypted);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine();
    Console.WriteLine($"✓ users.dat erstellt: {Path.GetFullPath(path)}");
    Console.WriteLine($"  Kunde:  {customer}");
    Console.WriteLine($"  Nutzer: {emails.Count}");
    foreach (var e in emails)
        Console.WriteLine($"    • {e}");
    Console.ResetColor();
}

void GenerateLicenseKey()
{
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine("─── Lizenzschlüssel generieren ─────────────────────────");
    Console.ResetColor();

    Console.Write("Kunden-E-Mail (appsettings.json → License:Email): ");
    var email = Console.ReadLine()?.Trim() ?? "";

    Console.Write("Ablaufdatum (YYYY-MM-DD, z.B. 2027-12-31): ");
    var expiry = Console.ReadLine()?.Trim() ?? "";

    var key = GenerateKey(email, expiry);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine();
    Console.WriteLine("appsettings.json Einträge:");
    Console.WriteLine($"  \"Email\":  \"{email.ToLower().Trim()}\",");
    Console.WriteLine($"  \"Key\":    \"{key}\",");
    Console.WriteLine($"  \"Expiry\": \"{expiry}\"");
    Console.ResetColor();
}

void ShowUserFile()
{
    Console.Write("Pfad zur users.dat (Enter = users.dat im aktuellen Verzeichnis): ");
    var path = Console.ReadLine()?.Trim();
    if (string.IsNullOrEmpty(path)) path = "users.dat";

    if (!File.Exists(path))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Datei nicht gefunden: {path}");
        Console.ResetColor();
        return;
    }

    try
    {
        var data = File.ReadAllBytes(path);
        var list = Decrypt(data);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Kunde:   {list.Customer}");
        Console.WriteLine($"Erstellt: {list.Created}");
        Console.WriteLine($"Nutzer:  {list.Users.Count}");
        foreach (var e in list.Users)
            Console.WriteLine($"  • {e}");
        Console.ResetColor();
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Fehler: {ex.Message}");
        Console.ResetColor();
    }
}

void EditAiSettings()
{
    Console.ForegroundColor = ConsoleColor.White;
    Console.WriteLine("─── KI-Einstellungen bearbeiten ─────────────────────────");
    Console.ResetColor();

    Console.Write("Pfad zur ai-settings.json (Enter = ai-settings.json im aktuellen Verzeichnis): ");
    var path = Console.ReadLine()?.Trim();
    if (string.IsNullOrEmpty(path)) path = "ai-settings.json";

    var current = new AiSettingsSection();
    if (File.Exists(path))
    {
        try
        {
            var loaded = JsonSerializer.Deserialize<AiSettingsFile>(File.ReadAllText(path));
            if (loaded?.AI != null) current = loaded.AI;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Warnung: Datei konnte nicht gelesen werden ({ex.Message}), starte mit leeren Werten.");
            Console.ResetColor();
        }
    }

    Console.WriteLine();
    Console.WriteLine("Aktuell:");
    Console.WriteLine($"  Provider: {current.Provider}");
    Console.WriteLine($"  API-Key:  {MaskKey(current.ApiKey)}");
    Console.WriteLine($"  Model:    {(current.Model.Length > 0 ? current.Model : "(Provider-Standard)")}");
    Console.WriteLine($"  Endpoint: {(current.Endpoint.Length > 0 ? current.Endpoint : "(Provider-Standard)")}");
    Console.WriteLine();

    var providers = new[] { "gemini", "anthropic", "openai", "groq", "ollama", "lmstudio" };
    Console.WriteLine("Provider wählen (Enter = unverändert lassen):");
    for (var i = 0; i < providers.Length; i++)
        Console.WriteLine($"  [{i + 1}]  {providers[i]}" + (providers[i] == current.Provider ? "  (aktuell)" : ""));
    Console.Write("Auswahl: ");
    var providerChoice = Console.ReadLine()?.Trim();
    var provider = current.Provider;
    if (int.TryParse(providerChoice, out var idx) && idx >= 1 && idx <= providers.Length)
        provider = providers[idx - 1];

    var isLocal = provider is "ollama" or "lmstudio";

    var apiKey = current.ApiKey;
    if (!isLocal)
    {
        Console.Write($"API-Key (Enter = unverändert, aktuell: {MaskKey(current.ApiKey)}): ");
        var keyInput = ReadMasked();
        if (!string.IsNullOrEmpty(keyInput)) apiKey = keyInput;
    }

    Console.Write("Model (Enter = unverändert): ");
    var modelInput = Console.ReadLine()?.Trim();
    var model = string.IsNullOrEmpty(modelInput) ? current.Model : modelInput;

    var endpoint = current.Endpoint;
    if (isLocal)
    {
        var def = provider == "lmstudio" ? "http://localhost:1234" : "http://localhost:11434";
        Console.Write($"Endpoint (Enter = unverändert, {(current.Endpoint.Length > 0 ? "aktuell: " + current.Endpoint : "Standard: " + def)}): ");
        var endpointInput = Console.ReadLine()?.Trim();
        if (!string.IsNullOrEmpty(endpointInput)) endpoint = endpointInput;
    }

    var result = new AiSettingsFile { AI = new AiSettingsSection { Provider = provider, ApiKey = apiKey, Model = model, Endpoint = endpoint } };
    File.WriteAllText(path, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine();
    Console.WriteLine($"✓ Gespeichert: {Path.GetFullPath(path)}");
    Console.WriteLine($"  Provider: {provider}");
    Console.WriteLine($"  API-Key:  {MaskKey(apiKey)}");
    Console.WriteLine($"  Model:    {(model.Length > 0 ? model : "(Provider-Standard)")}");
    Console.WriteLine($"  Endpoint: {(endpoint.Length > 0 ? endpoint : "(Provider-Standard)")}");
    Console.WriteLine("  Läuft tembu-server bereits, übernimmt er die Änderung sofort (kein Neustart nötig).");
    Console.ResetColor();
}

static string MaskKey(string key)
{
    if (string.IsNullOrEmpty(key)) return "(nicht gesetzt)";
    if (key.Length <= 8) return new string('•', key.Length);
    return key[..4] + new string('•', key.Length - 8) + key[^4..];
}

static string ReadMasked()
{
    if (Console.IsInputRedirected) return Console.ReadLine()?.Trim() ?? "";

    var sb = new StringBuilder();
    ConsoleKeyInfo key;
    while ((key = Console.ReadKey(intercept: true)).Key != ConsoleKey.Enter)
    {
        if (key.Key == ConsoleKey.Backspace)
        {
            if (sb.Length > 0) { sb.Length--; Console.Write("\b \b"); }
            continue;
        }
        if (!char.IsControl(key.KeyChar))
        {
            sb.Append(key.KeyChar);
            Console.Write('*');
        }
    }
    Console.WriteLine();
    return sb.ToString();
}

// ── Crypto ────────────────────────────────────────────────────────────────────

static byte[] DeriveKey()
    => SHA256.HashData(Encoding.UTF8.GetBytes(Secret));

static byte[] Encrypt(UserListData list)
{
    var json      = JsonSerializer.Serialize(list);
    var plaintext = Encoding.UTF8.GetBytes(json);
    var key       = DeriveKey();

    var nonce      = new byte[AesGcm.NonceByteSizes.MaxSize];
    var tag        = new byte[AesGcm.TagByteSizes.MaxSize];
    var ciphertext = new byte[plaintext.Length];

    RandomNumberGenerator.Fill(nonce);

    using var aes = new AesGcm(key, AesGcm.TagByteSizes.MaxSize);
    aes.Encrypt(nonce, plaintext, ciphertext, tag);

    var result = new byte[nonce.Length + tag.Length + ciphertext.Length];
    Buffer.BlockCopy(nonce,      0, result, 0,                         nonce.Length);
    Buffer.BlockCopy(tag,        0, result, nonce.Length,              tag.Length);
    Buffer.BlockCopy(ciphertext, 0, result, nonce.Length + tag.Length, ciphertext.Length);
    return result;
}

static UserListData Decrypt(byte[] data)
{
    var nonceSize  = AesGcm.NonceByteSizes.MaxSize;
    var tagSize    = AesGcm.TagByteSizes.MaxSize;

    var nonce      = data[..nonceSize];
    var tag        = data[nonceSize..(nonceSize + tagSize)];
    var ciphertext = data[(nonceSize + tagSize)..];
    var plaintext  = new byte[ciphertext.Length];

    using var aes = new AesGcm(DeriveKey(), AesGcm.TagByteSizes.MaxSize);
    aes.Decrypt(nonce, ciphertext, tag, plaintext);

    return JsonSerializer.Deserialize<UserListData>(plaintext)!;
}

static string GenerateKey(string email, string expiry)
{
    var input = $"{email.ToLower().Trim()}|{expiry}|{Secret}";
    using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(Secret));
    var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(input));
    var hex  = BitConverter.ToString(hash).Replace("-", "").ToUpper();
    return $"{Prefix}-{hex[..4]}-{hex[4..8]}-{hex[8..12]}-{hex[12..16]}";
}

// ── Modell ────────────────────────────────────────────────────────────────────

record UserListData
{
    public string Customer { get; init; } = "";
    public string Created  { get; init; } = "";
    public List<string> Users { get; init; } = [];
}

// Struktur muss zu tembu-server\Models\AiSettings.cs / Program.cs-Sektion "AI" passen
record AiSettingsFile
{
    public AiSettingsSection AI { get; init; } = new();
}

record AiSettingsSection
{
    public string Provider { get; init; } = "gemini";
    public string ApiKey   { get; init; } = "";
    public string Model    { get; init; } = "";
    public string Endpoint { get; init; } = "";
}
