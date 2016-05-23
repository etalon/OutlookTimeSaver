Imports log4net
Imports log4net.Repository.Hierarchy
Imports log4net.Repository.Hierarchy.logger

Public Class ThisAddIn

    Private WithEvents m_Inspectors As Outlook.Inspectors
    Private WithEvents m_Explorer As Outlook.Explorer

    Private m_MovedItemsUnreader As DeletedItemsHandler

    ''' <summary>
    ''' Wenn wir ein MailItem über den Explorer abgreifen wird seltsamerweise auch 
    ''' ein Inspector-Event mit einem ungültigen MailItem ausgelöst.
    ''' </summary>
    Private m_AllowNewMailItemsByInspector As Boolean = True

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        Try

            UnhandledExceptionHandler.Activate()

            Config.Load()
            DatabaseCreator.Init()

            With DirectCast(LogManager.GetRepository, Hierarchy)
                With .Root
                    If Config.My.DebugViewMode Then
                        .AddAppender(New LoggerFormAppender)
                    End If

                    If Config.My.LoggingEnabled Then
                        .AddAppender(MyRollingFileAppender)
                    End If

                    .Level = Core.Level.All
                End With
                .Configured = True

            End With

            m_Inspectors = Application.Inspectors
            m_Explorer = Application.ActiveExplorer

            MailItemHandler.PassOutlookApplication(Application)

            m_MovedItemsUnreader = New DeletedItemsHandler(Application)

            Log.Debug("OutlookTimeSaver-Addin started")

        Catch ex As Exception
            Log.Fatal("Schwerer Fehler bei Programmstart", ex)
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub m_Explorer_InlineResponse() Handles m_Explorer.InlineResponse

        Dim setSalutation As Boolean

        Try
            m_AllowNewMailItemsByInspector = False
            setSalutation = m_Explorer.CurrentFolder.EntryID <> Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).EntryID
            MailItemHandlerList.Add(TryCast(m_Explorer.ActiveInlineResponse, Outlook.MailItem), True, setSalutation)
        Finally
            m_AllowNewMailItemsByInspector = True
        End Try

    End Sub

    Private Sub inspectors_NewInspector(ByVal p_Inspector As Microsoft.Office.Interop.Outlook.Inspector) Handles m_Inspectors.NewInspector

        Log.Debug("NewInspector: " & p_Inspector.Caption)

        If Not m_AllowNewMailItemsByInspector Then
            Log.Debug("Keine neuen MailItems über den Inspector zur Zeit erlaubt.")
            Return
        End If

        Try

            Select Case True
                Case TypeOf p_Inspector.CurrentItem Is Outlook.MailItem
                    MailItemHandlerList.Add(TryCast(p_Inspector.CurrentItem, Outlook.MailItem), False, True)
                Case Else
                    Log.Debug("Inspector (" & p_Inspector.Caption & ") nicht aufgenommen")
            End Select

        Catch ex As Exception
            Log.Fatal("NewInspectorEvent", ex)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "NewInspectorEvent")
        End Try

    End Sub

End Class
