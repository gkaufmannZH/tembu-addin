namespace tembu_server.Models;

public class LicenseInfo
{
    public bool IsValid { get; set; }
    public string Email { get; set; } = "";
    public DateTime ExpiryDate { get; set; }
    public string ErrorMessage { get; set; } = "";

    public bool IsExpired => DateTime.Today > ExpiryDate;
    public int DaysRemaining => Math.Max(0, (ExpiryDate - DateTime.Today).Days);
}
