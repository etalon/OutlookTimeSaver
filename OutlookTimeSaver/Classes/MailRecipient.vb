Imports System.Net.Mail

Public Class MailRecipient

    Public Enum GenderEnum
        Male
        Female
    End Enum

    Private m_Email As MailAddress
    Private m_FirstName As String
    Private m_LastName As String
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

    Public Sub New(p_Email As MailAddress)

        Dim outlookContact As Outlook.ContactItem = Nothing

        m_Email = p_Email

        ' Email in Kontakten suchen
        If OutlookContacts.TryGetContact(m_Email.ToString, outlookContact) Then
            m_FirstName = outlookContact.FirstName
            m_LastName = outlookContact.LastName
        Else
            ' Email manuell splitten und versuchen Vor- und Nachnamen auszulesen
            resolveNameByEmail(p_Email.ToString)
        End If

        resolveGender()

    End Sub

    Private Sub resolveGender()

        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()

            Select Case db.ReadScalarDefault(Of String)("SELECT gender FROM firstname WHERE LOWER(name) = LOWER(@0)", "m", m_FirstName).ToLower
                Case "m"
                    m_Gender = GenderEnum.Male
                Case "w"
                    m_Gender = GenderEnum.Female
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

        ' TODO: Vorname in Namensdatenbank suchen.
        ' TODO: Standard-Anrede festlegen

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
