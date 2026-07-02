Public Class LicenseInfo

    Public Property IsValid As Boolean
    Public Property Email As String
    Public Property ExpiryDate As DateTime
    Public Property MaxUsers As Integer
    Public Property ErrorMessage As String

    Public ReadOnly Property IsExpired As Boolean
        Get
            Return DateTime.Today > ExpiryDate
        End Get
    End Property

    Public ReadOnly Property DaysRemaining As Integer
        Get
            Return Math.Max(0, (ExpiryDate - DateTime.Today).Days)
        End Get
    End Property

End Class
