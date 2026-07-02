using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using tembu_server.Models;

namespace tembu_server.Services;

public class UserListService
{
    // Gleicher Secret wie LicenseService — nie ändern
    private const string Secret = "tembu-2024-x7k9-secure";

    private readonly HashSet<string> _authorizedUsers = new(StringComparer.OrdinalIgnoreCase);
    private readonly ILogger<UserListService> _logger;
    private bool _fileExists = false;

    public UserListService(ILogger<UserListService> logger, IConfiguration config)
    {
        _logger = logger;
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

    // ── Verschlüsseln (für Generator-Tool) ────────────────────────────────
    public static byte[] Encrypt(UserList list)
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

        // Format: [nonce 12 bytes][tag 16 bytes][ciphertext]
        var result = new byte[nonce.Length + tag.Length + ciphertext.Length];
        Buffer.BlockCopy(nonce,       0, result, 0,                         nonce.Length);
        Buffer.BlockCopy(tag,         0, result, nonce.Length,              tag.Length);
        Buffer.BlockCopy(ciphertext,  0, result, nonce.Length + tag.Length, ciphertext.Length);
        return result;
    }

    // ── Entschlüsseln ─────────────────────────────────────────────────────
    private static string Decrypt(byte[] data)
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

    private static byte[] DeriveKey()
        => SHA256.HashData(Encoding.UTF8.GetBytes(Secret));
}
