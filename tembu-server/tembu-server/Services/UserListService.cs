using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using tembu_server.Models;

namespace tembu_server.Services;

public class UserListService
{
    // Gleicher Secret wie LicenseService — kommt aus License:Secret (appsettings.Local.json),
    // NICHT hartcodiert, damit er nicht im öffentlichen Repo landet.
    private readonly string _secret;

    private readonly HashSet<string> _authorizedUsers = new(StringComparer.OrdinalIgnoreCase);
    private readonly ILogger<UserListService> _logger;
    private bool _fileExists = false;

    public UserListService(ILogger<UserListService> logger, IConfiguration config)
    {
        _logger = logger;
        _secret = config["License:Secret"] ?? "";
        var path = config["UserList:Path"] ?? "users.dat";
        Load(path);
    }

    // Ist kein File vorhanden → alle Nutzer erlaubt (Solo-Betrieb)
    public bool IsAuthorized(string email)
    {
        if (!_fileExists) return true;
        return _authorizedUsers.Contains(email.Trim().ToLower());
    }

    public int UserCount => _authorizedUsers.Count;

    // ── File laden und entschlüsseln ───────────────────────────────────────
    private void Load(string path)
    {
        if (!File.Exists(path))
        {
            _logger.LogWarning("users.dat nicht gefunden — alle Nutzer erlaubt (Solo-Betrieb)");
            return;
        }

        try
        {
            var data = File.ReadAllBytes(path);
            var json = Decrypt(data);
            var list = JsonSerializer.Deserialize<UserList>(json) ?? new UserList();

            foreach (var email in list.Users)
                _authorizedUsers.Add(email.Trim().ToLower());

            _fileExists = true;
            _logger.LogInformation("users.dat geladen: {Count} Nutzer für '{Customer}'",
                _authorizedUsers.Count, list.Customer);
        }
        catch (Exception ex)
        {
            _logger.LogError("users.dat konnte nicht geladen werden: {Message}", ex.Message);
        }
    }

    // ── Entschlüsseln ─────────────────────────────────────────────────────
    private string Decrypt(byte[] data)
    {
        var nonceSize  = AesGcm.NonceByteSizes.MaxSize;
        var tagSize    = AesGcm.TagByteSizes.MaxSize;

        var nonce      = data[..nonceSize];
        var tag        = data[nonceSize..(nonceSize + tagSize)];
        var ciphertext = data[(nonceSize + tagSize)..];
        var plaintext  = new byte[ciphertext.Length];

        var key = DeriveKey();
        using var aes = new AesGcm(key, AesGcm.TagByteSizes.MaxSize);
        aes.Decrypt(nonce, ciphertext, tag, plaintext);

        return Encoding.UTF8.GetString(plaintext);
    }

    private byte[] DeriveKey()
        => SHA256.HashData(Encoding.UTF8.GetBytes(_secret));
}
