Imports System.Security.Cryptography
Imports System.Text

' Tembu Lizenz-Generator
' Nur lokal bei Ihnen verwenden — nie an Kunden weitergeben

Module Program

    Private Const SECRET As String = "tembu-2024-x7k9-secure"
    Private Const PREFIX As String = "TEMBU"

    Sub Main()
        Console.WriteLine("═══════════════════════════════")
        Console.WriteLine("  Tembu Lizenz-Generator")
        Console.WriteLine("═══════════════════════════════")
        Console.WriteLine()

        Console.Write("E-Mail des Kunden : ")
        Dim email As String = Console.ReadLine().Trim().ToLower()

        Console.Write("Ablauf (YYYY-MM-DD): ")
        Dim expiry As String = Console.ReadLine().Trim()

        Console.Write("Max. Nutzer       : ")
        Dim maxUsers As Integer = Integer.Parse(Console.ReadLine().Trim())

        Dim key As String = GenerateKey(email, expiry, maxUsers)

        Console.WriteLine()
        Console.WriteLine("───────────────────────────────")
        Console.WriteLine($"E-Mail  : {email}")
        Console.WriteLine($"Ablauf  : {expiry}")
        Console.WriteLine($"Nutzer  : {maxUsers}")
        Console.WriteLine($"Schlüssel: {key}")
        Console.WriteLine("───────────────────────────────")
        Console.WriteLine()
        Console.WriteLine("In appsettings.json beim Kunden:")
        Console.WriteLine($"""
  "License": {{
    "Email":    "{email}",
    "Key":      "{key}",
    "Expiry":   "{expiry}",
    "MaxUsers": "{maxUsers}"
  }}
""")
        Console.ReadLine()
    End Sub

    Private Function GenerateKey(email As String, expiry As String, maxUsers As Integer) As String
        Dim input As String = $"{email}|{expiry}|{maxUsers}|{SECRET}"
        Using hmac As New HMACSHA256(Encoding.UTF8.GetBytes(SECRET))
            Dim hash As Byte() = hmac.ComputeHash(Encoding.UTF8.GetBytes(input))
            Dim hex As String = BitConverter.ToString(hash).Replace("-", "").ToUpper()
            Return $"{PREFIX}-{hex.Substring(0, 4)}-{hex.Substring(4, 4)}-{hex.Substring(8, 4)}-{hex.Substring(12, 4)}"
        End Using
    End Function

End Module
