﻿Imports System.Net.Mail
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
    Private m_BodyFormat As Outlook.OlBodyFormat

    Private m_KnownPropertyChanges As New List(Of String)

    Public Shared Sub PassOutlookApplication(p_OutlookApplication As Outlook.Application)
        m_OutlookApplication = p_OutlookApplication
    End Sub

    Private ReadOnly Property isSalutationWritten As Boolean
        Get
            Return Not String.IsNullOrEmpty(m_LastSalutationWritten)
        End Get
    End Property

    Private ReadOnly Property salutationTableKey As String
        Get
            If m_Recipients.Count = 0 Then
                Throw New Exception("Es konnten keine Empfänger ermittelt werden.")
            End If

            Return Join(m_Recipients.Select(Function(x) x.EMailAsString).ToArray, ",")
        End Get
    End Property

    Public Sub New(p_MailItem As Outlook.MailItem)

        Log.Debug("New MailItem")
        m_MailItem = p_MailItem
        m_BodyFormat = m_MailItem.BodyFormat

    End Sub

    Public Sub SaveSalutationToReceipients()

        Dim salutation As String = getCurrentSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
            db.ExecuteNonQuery("INSERT OR REPLACE INTO salutation (recipients,text) VALUES (@0,@1);", salutationTableKey, salutation)
        End Using

        Log.Debug(String.Format("Anrede zu {0} wurde aktualisiert: {1}", salutationTableKey, salutation))

    End Sub

    Private Function getCurrentSalutation() As String

        Dim salutation As String = ""

        ' TODO: Anrede wird momentan immer aktualisiert
        With m_WordEditor.Application.Selection
            .Start = 0
            .End = .EndKey(WordEnums.WDUnits.wdLine, WordEnums.WDMovementType.wdExtend)
            .Start = 0
            salutation = .Text
        End With

        If String.IsNullOrEmpty(salutation) Then
            Log.Debug("Es konnte keine Anrede gelesen werden.")
            Return ""
        End If

        salutation = salutation.Trim
        Log.Debug("Gelesene erste Zeile: " & salutation)

        If Not salutation.EndsWith(",") AndAlso Not salutation.EndsWith(".") AndAlso Not salutation.EndsWith("!") Then
            Return "" ' Keine gültige Anrede gefunden...
        End If

        If Not VALID_SALUTATIONS.Any(Function(x) salutation.StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) Then
            Return "" ' Keine gültige Anrede gefunden...
        End If

        Log.Debug("Finale Anrede für Datenbank: " & salutation)

        Return salutation

    End Function

    Private Sub m_MailItem_Open() Handles m_MailItem.Open

        Log.Debug("MailItem_Open")

        m_WordEditor = DirectCast(m_MailItem.GetInspector.WordEditor, Word.Document)

        If String.IsNullOrEmpty(m_MailItem.To) Then
            m_IsNewMail = True
            Return ' Neue Mail
        End If

        m_AfterMailOpenThread = New Thread(AddressOf runAfterResponseMailOpenThread)
        m_AfterMailOpenThread.Start()

    End Sub

    Private Sub m_MailItem_Send() Handles m_MailItem.Send

        Log.Debug("Nachricht wird gesendet...")

        SaveSalutationToReceipients()
        MailItemHandlerList.Remove(Me)

    End Sub

    Private Sub runAfterResponseMailOpenThread()

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

            If Not m_IsNewMail Then
                Return
            End If

            m_KnownPropertyChanges.Add(Name.ToLower)

            Select Case Name.ToLower
                Case "to"
                    If m_KnownPropertyChanges.Contains("subject") Or m_MailItem.Subject.StartsWith("WG:") Then
                        setRecipientsAndSaluation()
                    End If
                Case "subject"
                    If m_KnownPropertyChanges.Contains("to") Then
                        setRecipientsAndSaluation()
                    End If
            End Select

        Catch ex As Exception
            Log.Fatal("MailItem_PropertyChange", ex)
        End Try

    End Sub

    Private Sub setRecipientsAndSaluation()

        m_KnownPropertyChanges.Clear()

        If Config.My.NoSalutationAtTopicStartsWith.Exists(Function(x) m_MailItem.Subject.StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) Then
            Log.Debug("Anrede wird nicht gesetzt, da Überschrift in der Ausschlussliste enthalten ist.")
            Return
        End If

        setRecipients()
        setSalutationByWordEditor()

    End Sub

    Private Sub setRecipients()

        If Not m_MailItem.Recipients.ResolveAll Then
            Log.Debug("ResolveAll failed")
        End If

        Log.Debug("SetRecipients.Count: " & m_MailItem.Recipients.Count & " / m_Recipients.Count: " & m_Recipients.Count & " / TO: " & m_MailItem.To)
        m_Recipients.Clear()

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

            m_Recipients.Add(New MailRecipient(rec))

        Next

    End Sub

    Private Sub setSalutationByWordEditor()

        Dim salutation As String = getAutomaticSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        If Not String.IsNullOrEmpty(m_LastSalutationWritten) Then
            With m_WordEditor.Application.Selection
                .Start = 0
                .End = m_LastSalutationWritten.Length + 2
                .Delete()
            End With
        End If

        With m_WordEditor.Application.Selection
            .Start = 0
            .InsertBefore(salutation & vbCrLf & vbCrLf)
            .Start = salutation.Length + 2
            .End = .Start
        End With

    End Sub

    Private Sub setSalutationByBody()

        Dim salutation As String = ""
        Dim newMailBody As String = ""
        Dim matches As MatchCollection

        Log.Debug("Anrede setzen...")

        salutation = getAutomaticSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        If Not String.IsNullOrEmpty(m_LastSalutationWritten) Then

            Log.Debug("Letzte Anrede überschreiben...")

            Select Case m_BodyFormat
                Case Outlook.OlBodyFormat.olFormatHTML
                    newMailBody = m_MailItem.HTMLBody.Replace(m_LastSalutationWritten, salutation)
                Case Else
                    newMailBody = m_MailItem.Body.Replace(m_LastSalutationWritten, salutation)

                    If Not newMailBody.Contains(", " & vbCrLf & vbCrLf) Then
                        newMailBody = Replace(newMailBody, ", " & vbCrLf, ", " & vbCrLf & vbCrLf)
                    End If

            End Select
        Else

            Log.Debug("Neue Anrede setzen...")

            Select Case m_BodyFormat
                Case Outlook.OlBodyFormat.olFormatHTML

                    newMailBody = m_MailItem.HTMLBody

                    matches = Regex.Matches(newMailBody, "<body [^<^>]*><div class=WordSection1><p class=MsoNormal><o:p>&nbsp;<\/o:p>", RegexOptions.None)
                    If matches.Count <> 1 Then
                        matches = Regex.Matches(newMailBody, "<body [^<^>]*><div class=WordSection1><p class=MsoNormal><span [^<^>]*><o:p>&nbsp;<\/o:p>", RegexOptions.None)

                        If matches.Count <> 1 Then
                            Log.Debug(newMailBody)
                            Throw New Exception("Stelle zum Einfügen der Anrede wurde nicht gefunden")
                        End If

                    End If

                    newMailBody = Replace(newMailBody, matches(0).ToString, Replace(matches(0).ToString, "&nbsp;", salutation & "<br><br>.<br>"))

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
                Log.Debug("Automatisch ermittelte Anrede: n.a - keine Empfänger")
                Return ""
            Case 1 To 2

                Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                    salutation = db.ReadScalarDefault(Of String)("SELECT text FROM salutation WHERE recipients = @0", "", salutationTableKey)
                End Using

                Log.Debug("Letzte Anrede aus Datenbank: " & salutation)

                If Not String.IsNullOrEmpty(salutation) Then
                    Return salutation
                End If

                For Each rec In m_Recipients
                    salutation &= rec.DefaultSalutation & ", "
                Next

            Case Else
                salutation = "Sehr geehrte Damen und Herren, "
        End Select

        Log.Debug("Automatisch ermittelte Anrede: " & salutation)

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
            Select Case m_BodyFormat
                Case Outlook.OlBodyFormat.olFormatHTML
                    .Start = p_Salutation.Length + 2
                Case Else
                    .Start = p_Salutation.Length + 2
            End Select
        End With

    End Sub

End Class
