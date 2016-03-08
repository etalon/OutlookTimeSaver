Imports System.Net.Mail

Public Class MailRecipient

    Public Enum GenderEnum
        Male
        Female
    End Enum

    Private m_OutlookRecipient As Outlook.Recipient
    Private m_Email As MailAddress
    Private m_FirstName As String = ""
    Private m_LastName As String = ""
    Private m_Gender As GenderEnum

    Public ReadOnly Property FirstName As String
        Get
            Return m_FirstName
        End Get
    End Property

    Public ReadOnly Property LastName As String
        Get
            Return m_LastName
        End Get
    End Property

    Public ReadOnly Property EMailAsString As String
        Get
            Return m_Email.ToString
        End Get
    End Property

    Public ReadOnly Property DefaultSalutation As String
        Get

            Dim salutation As String = "Hallo "

            Select Case m_Gender
                Case GenderEnum.Male
                    salutation &= "Herr"
                Case GenderEnum.Female
                    salutation &= "Frau"
            End Select

            salutation &= " " & m_LastName
            Return salutation

        End Get
    End Property

    Public Sub New(p_OutlookRecipient As Outlook.Recipient)

        Dim outlookContact As Outlook.ContactItem = Nothing

        If p_OutlookRecipient.AddressEntry.Type = "EX" Then
            m_Email = New MailAddress(p_OutlookRecipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
        Else
            m_Email = New MailAddress(p_OutlookRecipient.Address)
        End If

        m_OutlookRecipient = p_OutlookRecipient

        ' Email in Kontakten suchen
        If OutlookContacts.TryGetContact(m_Email.ToString, outlookContact) Then
            m_FirstName = outlookContact.FirstName
            m_LastName = outlookContact.LastName
        Else
            ' Email manuell splitten und versuchen Vor- und Nachnamen auszulesen
            resolveNameByEmail(m_Email.ToString)
        End If

        Log.Debug("Vorname: " & m_FirstName & "/ Nachname: " & m_LastName)

        If Not String.IsNullOrEmpty(m_FirstName) Then
            resolveGender()
        Else
            Log.Debug("Vorname konnte nicht ausgewertet werden (" & m_OutlookRecipient.Name & ")")
        End If

    End Sub

    Private Sub resolveGender()

        Dim value As String

        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()

            value = db.ReadScalarDefault(Of String)("SELECT gender FROM firstname WHERE LOWER(name) = LOWER(@0)", "m", m_FirstName)

            Select Case value
                Case "m", "M", ""
                    m_Gender = GenderEnum.Male
                Case "w", "W"
                    m_Gender = GenderEnum.Female
                Case Else
                    Throw New Exception("Ungültiger Wert: " & value)
            End Select

        End Using

    End Sub

    Private Sub resolveNameByEmail(p_Email As String)

        Dim user() As String

        user = m_Email.User.Split("."c)

        If user.Length <> 2 Then
            Return
        End If

        m_FirstName = GetUppercasedName(user(0))
        m_LastName = GetUppercasedName(user(1))

    End Sub

    Public Function GetUppercasedName(p_Name As String) As String

        Dim chars() As Char = p_Name.ToCharArray

        For i As Integer = 0 To chars.Length - 1
            Select Case True
                Case i = 0 ' Der erste Buchstabe
                    chars(i) = Char.ToUpper(chars(i))
                Case chars(i - 1) = "-"c ' Oder der zweite Nachname bei einem Doppelnamen
                    chars(i) = Char.ToUpper(chars(i))
            End Select
        Next

        Return New String(chars)

    End Function

End Class
