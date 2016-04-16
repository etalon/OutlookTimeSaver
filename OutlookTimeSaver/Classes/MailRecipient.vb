Imports System.Net.Mail

Public Class MailRecipient

#Region "Konstanten"

    Public Enum GenderEnum
        Male
        Female
    End Enum

#End Region

#Region "Member"

    Private m_OutlookRecipient As Outlook.Recipient
    Private m_Email As MailAddress
    Private m_FirstName As String = ""
    Private m_LastName As String = ""
    Private m_Gender As GenderEnum
    Private m_ExistsInDatabase As BoolSetEnum

#End Region

#Region "Properties"

    ''' <summary>
    ''' Bestimmt ob der Eintrag noch gültig ist
    ''' </summary>
    ''' <returns></returns>
    Public Property Valid As Boolean = True

    Public ReadOnly Property DisplayName As String
        Get
            Return m_Email.DisplayName
        End Get
    End Property

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

    Public ReadOnly Property Gender As String
        Get
            Select Case m_Gender
                Case GenderEnum.Female
                    Return "W"
                Case GenderEnum.Male
                    Return "M"
                Case Else
                    Throw New NotSupportedException
            End Select
        End Get
    End Property

    Public Function GetSalutation(ByRef p_IsFromDatabase As Boolean) As String

        Dim ret As String = lastSalutation
        If Not String.IsNullOrEmpty(ret) Then
            p_IsFromDatabase = True
            Return ret
        End If

        Return DefaultSalutation

    End Function

    Private ReadOnly Property lastSalutation As String
        Get

            Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                Return db.ReadScalarDefault(Of String)("SELECT salutation FROM recipient WHERE email = @0", "", m_Email.ToString)
            End Using

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

    Public ReadOnly Property GetUppercasedName(p_Name As String) As String
        Get
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
        End Get
    End Property

#End Region

#Region "Konstruktoren"

    Public Sub New(p_OutlookRecipient As Outlook.Recipient)

        Dim outlookContact As Outlook.ContactItem = Nothing

        m_OutlookRecipient = p_OutlookRecipient

        readMailAddress()

        Log.Debug("MailRecipient.New: " & m_Email.ToString & "/" & m_OutlookRecipient.Address)

        If ExistsInDatabase() Then
            Return ' Wenn wir schon eine Anrede haben, müssen wir nicht mehr Vornamen, Namen und Geschlecht auslesen
        End If

        resolveName()

        If Not String.IsNullOrEmpty(m_FirstName) Then
            resolveGender()
        Else
            Log.Debug("Vorname konnte nicht ausgewertet werden (" & m_OutlookRecipient.Name & ")")
        End If

    End Sub

#End Region

    Private Sub readMailAddress()

        Dim mail As String = ""

        If m_OutlookRecipient.AddressEntry.Type = "EX" Then

            Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()

                mail = db.ReadScalarDefault(Of String)("SELECT email FROM exchangeaddress WHERE address = @0", "", m_OutlookRecipient.Address)

                If String.IsNullOrEmpty(mail) Then
                    mail = m_OutlookRecipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    db.ExecuteNonQuery("INSERT INTO exchangeaddress (address, email) VALUES (@0, @1)", m_OutlookRecipient.Address, mail)
                End If

                m_Email = New MailAddress(mail)

            End Using
        Else
            m_Email = New MailAddress(m_OutlookRecipient.Address)
        End If

    End Sub

    Private Sub resolveName()

        Dim firstNameFromEmail As String = ""
        Dim lastNameFromEmail As String = ""
        Dim firstNameFromDisplayName As String = ""
        Dim lastNameFromDisplayName As String = ""

        Dim resolvedByEmail, resolvedByDisplayName As Boolean

        resolvedByEmail = resolveNameByEmail(firstNameFromEmail, lastNameFromEmail)
        resolvedByDisplayName = resolveByDisplayName(firstNameFromDisplayName, lastNameFromDisplayName)

        Select Case True
            Case resolvedByEmail
                m_FirstName = firstNameFromEmail
                m_LastName = lastNameFromEmail
            Case resolvedByDisplayName
                m_FirstName = firstNameFromDisplayName
                m_LastName = lastNameFromDisplayName
        End Select

        If Not String.IsNullOrEmpty(lastNameFromDisplayName) AndAlso lastNameFromEmail <> lastNameFromDisplayName Then
            If lastNameFromEmail.SameText(replaceUmlauts(lastNameFromDisplayName)) Then
                m_LastName = lastNameFromDisplayName
            End If
        End If

        Log.Debug("Vorname: " & m_FirstName & "/ Nachname: " & m_LastName)

    End Sub

    Private Function replaceUmlauts(ByVal p_Value As String) As String

        p_Value = Replace(p_Value, "ä", "ae")
        p_Value = Replace(p_Value, "ö", "oe")
        p_Value = Replace(p_Value, "ü", "ue")
        p_Value = Replace(p_Value, "ß", "ss")

        Return p_Value

    End Function

    Public Function ExistsInDatabase() As Boolean

        If m_ExistsInDatabase = BoolSetEnum.NotSet Then
            Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()

                Select Case db.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM recipient WHERE email = @0", m_Email.ToString) > 0
                    Case True
                        m_ExistsInDatabase = BoolSetEnum.True
                    Case False
                        m_ExistsInDatabase = BoolSetEnum.False
                End Select

                Return ExistsInDatabase

            End Using
        Else
            Return m_ExistsInDatabase = BoolSetEnum.True
        End If

    End Function

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

    Private Function resolveNameByEmail(ByRef p_FirstName As String, ByRef p_LastName As String) As Boolean

        Dim user() As String = m_Email.User.Split("."c)

        If user.Length <> 2 Then
            Return False
        End If

        p_FirstName = GetUppercasedName(user(0))
        p_LastName = GetUppercasedName(user(1))
        Return True

    End Function

    Private Function resolveByDisplayName(ByRef p_FirstName As String, ByRef p_LastName As String) As Boolean

        Dim user() As String = m_OutlookRecipient.Name.Split(" "c)

        If user.Length <> 2 Then
            Return False
        End If

        p_FirstName = GetUppercasedName(user(0))
        p_LastName = GetUppercasedName(user(1))
        Return True

    End Function

End Class
