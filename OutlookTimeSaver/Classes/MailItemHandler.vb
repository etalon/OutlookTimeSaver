Imports System.Net.Mail
Imports System.Threading
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class MailItemHandler

    Private WithEvents m_MailItem As Outlook.MailItem
    Private Shared m_OutlookApplication As Outlook.Application

    Private m_Inspector As Outlook.Inspector

    Private m_MailToLine As String = ""
    Private m_Recipients As New List(Of MailRecipient)

    Private m_WordEditor As Word.Document

    Private m_AfterMailOpenThread As Thread
    Private m_IsNewMail As Boolean
    Private m_LastSalutationWritten As String

    Public Shared Sub PassOutlookApplication(p_OutlookApplication As Outlook.Application)
        m_OutlookApplication = p_OutlookApplication
    End Sub

    Public ReadOnly Property UniqueId As String
        Get
            Return m_MailItem.ConversationIndex
        End Get
    End Property

    Private ReadOnly Property salutationTableKey As String
        Get
            Return Join(m_Recipients.Select(Function(x) x.EMailAsString).ToArray, ",")
        End Get
    End Property

    Public Sub New(p_MailItem As Outlook.MailItem)

        Log.Debug("New MailItem")
        m_MailItem = p_MailItem

    End Sub

    Public Sub SaveSalutationToReceipients()

        Dim salutation As String = getCurrentSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
            db.ExecuteNonQuery("INSERT OR REPLACE INTO salutation (recipients,text) VALUES (@0,@1);", salutationTableKey, salutation)
        End Using

    End Sub

    Private Function getCurrentSalutation() As String

        Dim salutation As String = Split(m_MailItem.Body, vbCrLf, 2)(0).Trim

        If Not salutation.EndsWith(",") AndAlso Not salutation.EndsWith(".") AndAlso Not salutation.EndsWith("!") Then
            Return "" ' Keine gültige Anrede gefunden...
        End If

        If Not VALID_SALUTATIONS.Any(Function(x) salutation.StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) Then
            Return "" ' Keine gültige Anrede gefunden...
        End If

        Return salutation

    End Function

    Private Sub m_MailItem_Open() Handles m_MailItem.Open

        Log.Debug("MailItem_Open")

        m_WordEditor = DirectCast(m_MailItem.GetInspector.WordEditor, Word.Document)

        If String.IsNullOrEmpty(m_MailItem.To) Then
            m_IsNewMail = True
            Return ' Neue Mail
        End If

        m_AfterMailOpenThread = New Thread(AddressOf runAfterMailOpenThread)
        m_AfterMailOpenThread.Start()

    End Sub

    Private Sub runAfterMailOpenThread()

        Try
            While m_OutlookApplication.ActiveInspector Is Nothing
                Thread.Sleep(50)
            End While

            Log.Debug("ActiveInspector ist nicht mehr nothing")

            setRecipientsAndSaluation()

        Catch ex As Exception
            Log.Fatal("runAfterMailOpenThread", ex)
        End Try

    End Sub

    Private Sub m_MailItem_PropertyChange(Name As String) Handles m_MailItem.PropertyChange

        Try

            Log.Debug("MailItem_PropertyChange: " & Name)

            Select Case Name.ToLower
                Case "to"
                    ' Aktuell nichts machen
                Case "subject"

                    m_MailItem.Save()

                    If m_IsNewMail Then
                        setRecipientsAndSaluation()
                    End If
            End Select

        Catch ex As Exception
            Log.Fatal("MailItem_PropertyChange", ex)
        End Try

    End Sub

    Private Sub setRecipientsAndSaluation()

        setRecipients()
        setSalutation()

    End Sub

    Private Sub setRecipients()

        Dim tmpAddr As String

        Log.Debug("SetRecipients.Count: " & m_MailItem.Recipients.Count)
        m_Recipients.Clear()

        m_MailItem.Recipients.ResolveAll()

        For Each rec As Outlook.Recipient In m_MailItem.Recipients

            If rec.Type <> OutlookRecipientType.To Then
                Continue For
            End If

            If Not rec.Resolved Then
                rec.Resolve()
            End If

            If String.IsNullOrEmpty(rec.Address) Then
                Continue For
            End If

            If rec.AddressEntry.Type = "EX" Then
                tmpAddr = rec.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            Else
                tmpAddr = rec.Address
            End If

            m_Recipients.Add(New MailRecipient(New MailAddress(tmpAddr)))

        Next

    End Sub

    Private Sub setSalutation()

        Dim salutation As String = ""
        Dim newMailBody As String = ""

        salutation = getAutomaticSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        If Not String.IsNullOrEmpty(m_LastSalutationWritten) Then
            Select Case m_MailItem.BodyFormat
                Case Outlook.OlBodyFormat.olFormatHTML
                    newMailBody = m_MailItem.HTMLBody.Replace(m_LastSalutationWritten, salutation)
                Case Else
                    newMailBody = m_MailItem.Body.Replace(m_LastSalutationWritten, salutation)

                    If Not newMailBody.Contains(", " & vbCrLf & vbCrLf) Then
                        newMailBody = Replace(newMailBody, ", " & vbCrLf, ", " & vbCrLf & vbCrLf)
                    End If

            End Select
        Else
            Select Case m_MailItem.BodyFormat
                Case Outlook.OlBodyFormat.olFormatHTML

                    newMailBody = m_MailItem.HTMLBody.Replace("<div class=WordSection1><p class=MsoNormal><o:p>&nbsp;</o:p>", "<!-- salutation_start --><div class=WordSection1><p class=MsoNormal><o:p>" & salutation & "</o:p><p class=MsoNormal><o:p></o:p></p><p class=MsoNormal><o:p><!-- cursor --></o:p></p>")

                Case Else
                    If Not String.IsNullOrEmpty(m_MailItem.Body) AndAlso m_MailItem.Body.Contains(vbCrLf & vbCrLf) Then
                        newMailBody = salutation & vbCrLf & vbCrLf

                        If Not m_IsNewMail Then
                            newMailBody &= vbCrLf & vbCrLf
                        End If

                        newMailBody &= m_MailItem.Body.Substring(m_MailItem.Body.IndexOf(vbCrLf & vbCrLf) + 4)
                    Else
                        newMailBody = salutation & vbCrLf & vbCrLf
                        newMailBody &= m_MailItem.Body
                    End If
            End Select
        End If

        m_LastSalutationWritten = salutation

        Select Case m_MailItem.BodyFormat
            Case Outlook.OlBodyFormat.olFormatHTML
                m_MailItem.HTMLBody = newMailBody
            Case Else
                m_MailItem.Body = newMailBody
        End Select

        setBodyCursorPosition(newMailBody, salutation)

    End Sub

    Private Function getAutomaticSalutation() As String

        Dim salutation As String = ""

        Log.Debug("Anzahl Empfänger: " & m_Recipients.Count)

        Select Case m_Recipients.Count
            Case 0
                Return ""
            Case 1 To 2

                Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                    salutation = db.ReadScalarDefault(Of String)("SELECT text FROM salutation WHERE recipients = @0", "", salutationTableKey)
                End Using

                If Not String.IsNullOrEmpty(salutation) Then
                    Return salutation
                End If

                ' TODO: Hier noch über die Vornamen-Datenbank gehen
                For Each rec In m_Recipients
                    salutation &= rec.DefaultSalutation & ", "
                Next

            Case Else
                salutation = "Sehr geehrte Damen und Herren, "
        End Select

        Return salutation

    End Function

    Private Function tryGetMailAddressFromString(ByVal p_MailAddressString As String, ByRef p_MailAddressObject As MailAddress) As Boolean

        Try
            Dim m As Match = Regex.Match(p_MailAddressString, "\((.*?)\)", RegexOptions.None)

            If m.Captures.Count = 1 Then
                p_MailAddressString = m.Captures.Item(0).Value.TrimStart("("c).TrimEnd(")"c)
            End If

            p_MailAddressObject = New MailAddress(p_MailAddressString)
            Return True
        Catch
            Return False
        End Try

    End Function

    Private Sub setBodyCursorPosition(p_NewBody As String, p_Salutation As String)

        With m_WordEditor.Application.Selection
            .Start = p_Salutation.Length + 2
        End With

    End Sub

End Class
