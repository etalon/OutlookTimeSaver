﻿Imports System.Net.Mail
Imports System.Threading
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class MailItemHandler
    Implements IDisposable

    Private Shared m_OutlookApplication As Outlook.Application
    Private Shared m_MailItemSaveSyncLock As New Object

    Private m_Inspector As Outlook.Inspector

    Private WithEvents m_MailItem As Outlook.MailItem
    Private WithEvents m_AfterResponseMailOpenTimer As Windows.Forms.Timer
    Private WithEvents m_MailItemSaveTimer As New Windows.Forms.Timer

    Private m_Recipients As New List(Of MailRecipient)
    Private m_WordEditor As Word.Document

    Private m_IsNewMail As Boolean
    Private m_IsInlineRespone As Boolean
    Private m_IsOpenedFromDrafts As Boolean
    Private m_IsItemSent As Boolean

    Private m_LastSalutationWritten As String
    Private m_KnownPropertyChanges As New HashSet(Of String)
    Private m_SalutationFromDatabase As String = ""
    Private m_ReceivedTime As Date


#Region "Properties"

    Public Property EntryId As String

    Private ReadOnly Property isSalutationWritten As Boolean
        Get
            Return Not String.IsNullOrEmpty(m_LastSalutationWritten)
        End Get
    End Property

    Private ReadOnly Property isForwardedMessage As Boolean
        Get
            Return Not String.IsNullOrEmpty(m_MailItem.Subject) AndAlso m_MailItem.Subject.StartsWith("WG:")
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

    Public ReadOnly Property HasManuallyChanged As Boolean
        Get
            Return m_ReceivedTime <> m_MailItem.ReceivedTime
        End Get
    End Property

#End Region

#Region "Constructor"

    Public Sub New(p_MailItem As Outlook.MailItem, p_IsInlineResponse As Boolean, p_IsOpenedFromDrafts As Boolean)

        Log.Debug(String.Format("New MailItem (IsInlineRespone = {0}, OpenedFromDrafts = {1}", p_IsInlineResponse, p_IsOpenedFromDrafts))

        m_MailItem = p_MailItem
        m_IsInlineRespone = p_IsInlineResponse
        m_IsOpenedFromDrafts = p_IsOpenedFromDrafts

        If m_IsInlineRespone Then
            m_MailItem_Open()
        End If

    End Sub

#End Region

#Region "Events"

    Private Sub m_MailItem_Open() Handles m_MailItem.Open

        Log.Debug("MailItem_Open - " & m_MailItem.EntryID)

        m_WordEditor = DirectCast(m_MailItem.GetInspector.WordEditor, Word.Document)

        If String.IsNullOrEmpty(m_MailItem.To) Then
            m_IsNewMail = True
            Return ' Neue Mail
        End If

        m_AfterResponseMailOpenTimer = New Windows.Forms.Timer
        m_AfterResponseMailOpenTimer.Interval = 1
        m_AfterResponseMailOpenTimer.Enabled = True

    End Sub

    Private Sub m_MailItem_Close(ByRef cancel As Boolean) Handles m_MailItem.Close

        Log.Debug("MailItem_Close")

        If m_IsItemSent Or m_IsOpenedFromDrafts Then
            ' Die Mail nur aus der Liste entfernen, aber nicht löschen. Entwurf wird also beibehalten
            Me.Dispose()
        Else
            Log.Debug("Löschung der Mail wird in Kürze überprüft")
            MailDeleter.Add(m_MailItem)
        End If

    End Sub

    Private Sub m_MailItem_Unload() Handles m_MailItem.Unload

        Log.Debug("MailItem_Unload")

    End Sub

    Private Sub m_MailItem_Send() Handles m_MailItem.Send

        Log.Debug("Nachricht wird gesendet...")
        m_IsItemSent = True

        SaveSalutationToReceipients()
        MailItemHandlerList.Remove(Me)

    End Sub

    ''' <summary>
    ''' Achtung: Innerhalb dieser Funktion darf kein "Save" ausgeführt werden. Dies zerstört die Recipients-Collection
    ''' </summary>
    ''' <param name="Name"></param>
    Private Sub m_MailItem_PropertyChange(Name As String) Handles m_MailItem.PropertyChange

        SyncLock m_MailItemSaveSyncLock

            Try

                If String.IsNullOrEmpty(Name) Then
                    Log.Debug("PropertyChange.Name ist leer")
                    Exit Sub
                End If

                Log.Debug("MailItem_PropertyChange: " & Name)

                If Not m_KnownPropertyChanges.Contains(Name.ToLower) Then
                    m_KnownPropertyChanges.Add(Name.ToLower)
                End If

                Select Case Name.ToLower
                    Case "to"

                    Case "bcc"

                        If m_IsNewMail AndAlso Not m_KnownPropertyChanges.Contains("subject") AndAlso Not isForwardedMessage Then
                            ' Wenn wir eine neue Mail haben und der Betreff wurde noch nicht gesetzt, müssen wir auch noch keine Anrede setzen
                            Return
                        End If

                        m_MailItemSaveTimer.Interval = 1
                        m_MailItemSaveTimer.Enabled = True

                    Case "subject"

                        If Not String.IsNullOrEmpty(m_LastSalutationWritten) Then
                            ' Die Anrede setzen wir nur ein einziges Mal zu Beginn
                            Return
                        End If

                        If Not m_IsNewMail Then
                            Return
                        End If

                        If Not m_KnownPropertyChanges.Contains("bcc") Then
                            Return
                        End If

                        m_MailItemSaveTimer.Interval = 1
                        m_MailItemSaveTimer.Enabled = True


                End Select

            Catch ex As Exception
                Log.Fatal("MailItem_PropertyChange", ex)
            End Try

        End SyncLock

    End Sub

    Private Sub SaveTimerTicke() Handles m_MailItemSaveTimer.Tick

        SyncLock m_MailItemSaveSyncLock
            m_MailItemSaveTimer.Enabled = False
            m_MailItem.Save()
            m_ReceivedTime = m_MailItem.ReceivedTime
            EntryId = m_MailItem.EntryID
        End SyncLock

    End Sub

    Private Sub m_MailItem_BeforeCheckNames() Handles m_MailItem.BeforeCheckNames

        Log.Debug("MailItem_BeforeCheckNames")

        If m_MailItem.UserProperties.Find("HardDelete") IsNot Nothing Then
            ' Mail ist bereits zum Löschen vorgemerkt und darf nicht mehr verändert werden
            Return
        End If

        setRecipientsAndSaluation()

    End Sub

    Private Sub runAfterResponseMailOpenTimer() Handles m_AfterResponseMailOpenTimer.Tick

        m_AfterResponseMailOpenTimer.Enabled = False

        Try
            If Not m_IsInlineRespone Then
                While m_OutlookApplication.ActiveInspector Is Nothing
                    Thread.Sleep(50)
                End While

                Log.Debug("ActiveInspector ist nicht mehr Nothing")
            End If

            setRecipientsAndSaluation()

        Catch ex As Exception
            Log.Fatal("runAfterMailOpenThread", ex)
        End Try

    End Sub

#End Region

#Region "Methods"

    Public Shared Sub PassOutlookApplication(p_OutlookApplication As Outlook.Application)
        m_OutlookApplication = p_OutlookApplication
    End Sub

    Public Sub SaveSalutationToReceipients()

        Dim salutation As String = getCurrentSalutation()

        If String.IsNullOrEmpty(salutation) Then
            Return
        End If

        If Not salutation.SameText(m_SalutationFromDatabase) Then

            Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                db.ExecuteNonQuery("INSERT Or REPLACE INTO recipient (email, salutation, mailcount) VALUES (@0,@1,@2);", salutationTableKey, salutation, 0)
            End Using

            Log.Debug(String.Format("Anrede zu {0} wurde aktualisiert: {1}", salutationTableKey, salutation))

        End If

        ' TODO: An dieser Stelle müssten wir eigentlich überprüfen ob der Vorname oder Nachname so übernommen wurde und es dann aktualisieren.
        ' Nur so könnte man aus der Datenbank heraus lernen, aber vielleicht ist es auch unnötig.
        If m_Recipients.Count = 1 Then
            With m_Recipients.First
                Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                    db.ExecuteNonQuery("UPDATE recipient SET firstname = @0, lastname = @1, gender = @2, displayname = @3 WHERE email = @4;", .FirstName, .LastName, .Gender, .DisplayName, .EMailAsString)
                End Using
            End With
        End If

        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
            db.ExecuteNonQuery("UPDATE recipient SET mailcount = mailcount + 1 WHERE email = @0;", salutationTableKey)
        End Using

    End Sub

    Private Function getCurrentSalutation() As String

        Dim salutation As String = ""

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

    Private Sub setRecipientsAndSaluation()

        Dim haveRecipientsChanged As Boolean

        If m_IsOpenedFromDrafts Then
            Log.Debug("Anrede nicht setzen...")
            Return
        End If

        If Config.My.NoSalutationAtTopicStartsWith.Exists(Function(x) m_MailItem.Subject.StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) Then
            Log.Debug("Anrede wird nicht gesetzt, da Überschrift in der Ausschlussliste enthalten ist.")
            Return
        End If

        setRecipients(haveRecipientsChanged)

        If haveRecipientsChanged Then
            Log.Debug("Empfänger haben sich geändert")
            setSalutationByWordEditor()
        End If

    End Sub

    Private Sub setRecipients(ByRef p_HaveRecipientsChanged As Boolean)

        Dim newRecipient As MailRecipient
        Dim initialRecipientCount As Integer

        Log.Debug("SetRecipients.Count: " & m_MailItem.Recipients.Count & " / m_Recipients.Count: " & m_Recipients.Count) ' & " / TO: " & m_MailItem.To)

        For Each rec In m_Recipients
            rec.Valid = False
        Next

        initialRecipientCount = m_Recipients.Count

        For Each rec As Outlook.Recipient In m_MailItem.Recipients

            If rec.Type <> OutlookRecipientType.To Then
                Continue For
            End If

            If String.IsNullOrEmpty(rec.Address) Then
                Continue For
            End If

            newRecipient = New MailRecipient(rec)

            m_Recipients.RemoveAll(Function(x) x.EMailAsString = newRecipient.EMailAsString)
            m_Recipients.Add(newRecipient)

        Next

        p_HaveRecipientsChanged = initialRecipientCount <> m_Recipients.Count OrElse m_Recipients.Any(Function(x) Not x.Valid)
        m_Recipients.RemoveAll(Function(x) Not x.Valid)

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

        m_LastSalutationWritten = salutation

    End Sub

    Private Function getAutomaticSalutation() As String

        Dim salutation As String = ""
        Dim isFromDatabase As Boolean

        Log.Debug("Anzahl Empfänger: " & m_Recipients.Count)

        Select Case m_Recipients.Count
            Case 0
                Log.Debug("Automatisch ermittelte Anrede: n.a - keine Empfänger")
                Return ""
            Case 1

                salutation = m_Recipients(0).GetSalutation(isFromDatabase)
                If isFromDatabase Then
                    m_SalutationFromDatabase = salutation
                Else
                    ' Bei der Default-Anrede hängen wir noch das Komma an
                    salutation &= ", "
                End If

            Case 2

                Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
                    salutation = db.ReadScalarDefault(Of String)("SELECT salutation FROM recipient WHERE email = @0", "", salutationTableKey)
                End Using

                Log.Debug("Letzte Anrede aus Datenbank: " & salutation)

                If Not String.IsNullOrEmpty(salutation) Then
                    m_SalutationFromDatabase = salutation
                Else
                    salutation = Join(m_Recipients.Select(Function(x) x.DefaultSalutation).ToArray, ", ")
                    salutation &= ", "
                End If

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

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then

                If m_AfterResponseMailOpenTimer IsNot Nothing Then m_AfterResponseMailOpenTimer.Dispose()
                If m_MailItemSaveTimer IsNot Nothing Then m_MailItemSaveTimer.Dispose()
                MailItemHandlerList.Remove(Me)

            End If

            If m_MailItem IsNot Nothing Then
                Marshal.ReleaseComObject(m_MailItem)
                m_MailItem = Nothing
            End If

        End If
        disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region

End Class
