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
Console.WriteLine("║     TEMBU USER TOOL  v2.0            ║");
Console.WriteLine("╚══════════════════════════════════════╝");
Console.ResetColor();
Console.WriteLine();

// ── Modus wählen ─────────────────────────────────────────────────────────────
Console.WriteLine("Was möchten Sie tun?");
Console.WriteLine("  [1]  Neue users.dat erstellen");
Console.WriteLine("  [2]  Lizenzschlüssel generieren");
Console.WriteLine("  [3]  Bestehende users.dat anzeigen");
Console.Write("\nAuswahl: ");
var mode = Console.ReadLine()?.Trim();
Console.WriteLine();

switch (mode)
{
    case "1": CreateUserFile(); break;
    case "2": GenerateLicenseKey(); break;
    case "3": ShowUserFile(); break;
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
