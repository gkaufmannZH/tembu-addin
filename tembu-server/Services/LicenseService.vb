Imports System.Security.Cryptography
Imports System.Text

Public Class LicenseService

    ' Geheimnis — nur in Ihrem kompilierten Binary vorhanden
    ' VOR Produktionseinsatz ändern, danach NIE mehr ändern
    Private Const SECRET As String = "tembu-2024-x7k9-secure"
    Private Const PREFIX As String = "TEMBU"

    ' ── Lizenz prüfen ─────────────────────────────────────────────────────────
    Public Shared Function Validate(email As String,
                                    licenseKey As String,
                                    expiry As String,
                                    maxUsers As Integer) As LicenseInfo
        Dim info As New LicenseInfo()

        ' Eingaben prüfen
        If String.IsNullOrWhiteSpace(email) OrElse String.IsNullOrWhiteSpace(licenseKey) Then
            info.IsValid = False
            info.ErrorMessage = "E-Mail oder Lizenzschlüssel fehlt in appsettings.json."
            Return info
        End If

        ' Schlüssel prüfen
        Dim expected As String = GenerateKey(email, expiry, maxUsers)
        If Not String.Equals(licenseKey.Trim(), expected, StringComparison.OrdinalIgnoreCase) Then
            info.IsValid = False
            info.ErrorMessage = "Ungültiger Lizenzschlüssel. Bitte kontaktieren Sie support@tembu.ch."
            Return info
        End If

        ' Ablaufdatum prüfen
        Dim expiryDate As DateTime
        If Not DateTime.TryParse(expiry, expiryDate) Then
            info.IsValid = False
            info.ErrorMessage = "Ungültiges Ablaufdatum in appsettings.json (Format: YYYY-MM-DD)."
            Return info
        End If

        If DateTime.Today > expiryDate Then
            info.IsValid = False
            info.ErrorMessage = $"Lizenz abgelaufen am {expiryDate:dd.MM.yyyy}. Bitte erneuern Sie Ihre Lizenz."
            Return info
        End If

        ' Alles gültig
        info.IsValid = True
        info.Email = email.ToLower().Trim()
        info.ExpiryDate = expiryDate
        info.MaxUsers = maxUsers
        Return info
    End Function

    ' ── Schlüssel generieren (nur in Ihrem Generator-Tool verwenden) ───────────
    Public Shared Function GenerateKey(email As String,
                                       expiry As String,
                                       maxUsers As Integer) As String
        Dim input As String = $"{email.ToLower().Trim()}|{expiry}|{maxUsers}|{SECRET}"
        Using hmac As New HMACSHA256(Encoding.UTF8.GetBytes(SECRET))
            Dim hash As Byte() = hmac.ComputeHash(Encoding.UTF8.GetBytes(input))
            Dim hex As String = BitConverter.ToString(hash).Replace("-", "").ToUpper()
            Return $"{PREFIX}-{hex.Substring(0, 4)}-{hex.Substring(4, 4)}-{hex.Substring(8, 4)}-{hex.Substring(12, 4)}"
        End Using
    End Function

End Class
