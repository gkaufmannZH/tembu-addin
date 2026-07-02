using System.Security.Cryptography;
using System.Text;
using tembu_server.Models;

namespace tembu_server.Services;

public class LicenseService
{
    private const string Prefix = "TEMBU";

    public static LicenseInfo Validate(string email, string licenseKey, string expiry, string secret)
    {
        var info = new LicenseInfo();

        if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(licenseKey))
        {
            info.IsValid = false;
            info.ErrorMessage = "E-Mail oder Lizenzschlüssel fehlt in appsettings.json.";
            return info;
        }

        if (string.IsNullOrWhiteSpace(secret))
        {
            info.IsValid = false;
            info.ErrorMessage = "License:Secret fehlt (appsettings.Local.json). Ohne dieses Secret kann kein Lizenzschlüssel geprüft werden.";
            return info;
        }

        var expected = GenerateKey(email, expiry, secret);
        if (!string.Equals(licenseKey.Trim(), expected, StringComparison.OrdinalIgnoreCase))
        {
            info.IsValid = false;
            info.ErrorMessage = "Ungültiger Lizenzschlüssel. Bitte kontaktieren Sie support@tembu.ch.";
            return info;
        }

        if (!DateTime.TryParse(expiry, out var expiryDate))
        {
            info.IsValid = false;
            info.ErrorMessage = "Ungültiges Ablaufdatum in appsettings.json (Format: YYYY-MM-DD).";
            return info;
        }

        if (DateTime.Today > expiryDate)
        {
            info.IsValid = false;
            info.ErrorMessage = $"Lizenz abgelaufen am {expiryDate:dd.MM.yyyy}. Bitte erneuern Sie Ihre Lizenz.";
            return info;
        }

        info.IsValid = true;
        info.Email = email.ToLower().Trim();
        info.ExpiryDate = expiryDate;
        return info;
    }

    public static string GenerateKey(string email, string expiry, string secret)
    {
        var input = $"{email.ToLower().Trim()}|{expiry}|{secret}";
        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
        var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(input));
        var hex = BitConverter.ToString(hash).Replace("-", "").ToUpper();
        return $"{Prefix}-{hex[..4]}-{hex[4..8]}-{hex[8..12]}-{hex[12..16]}";
    }
}
