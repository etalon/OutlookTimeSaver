Imports log4net
Imports log4net.Repository.Hierarchy
Imports log4net.Repository.Hierarchy.logger

Public Class ThisAddIn

    Private WithEvents m_Inspectors As Outlook.Inspectors
    Private m_MovedItemsUnreader As DeletedItemsHandler

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

            Log.Debug("Started")

            m_Inspectors = Application.Inspectors

            MailItemHandler.PassOutlookApplication(Application)

            m_MovedItemsUnreader = New DeletedItemsHandler(Application)
            OutlookContacts.ReadContacts(Application)

        Catch ex As Exception
            Log.Fatal("Schwerer Fehler bei Programmstart", ex)
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub inspectors_NewInspector(ByVal p_Inspector As Microsoft.Office.Interop.Outlook.Inspector) Handles m_Inspectors.NewInspector

        Log.Debug("NewInspector: " & p_Inspector.Caption)

        Try

            Select Case True
                Case TypeOf p_Inspector.CurrentItem Is Outlook.MailItem
                    MailItemHandlerList.Add(TryCast(p_Inspector.CurrentItem, Outlook.MailItem))
                Case Else
                    Log.Debug("Inspector (" & p_Inspector.Caption & ") nicht aufgenommen")
            End Select

        Catch ex As Exception
            Log.Fatal("NewInspectorEvent", ex)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "NewInspectorEvent")
        End Try

    End Sub

    Private Sub inspector_ItemSend(p_Object As Object, ByRef p_Cancel As Boolean) Handles Application.ItemSend

        Dim myMailItemHandler As MailItemHandler

        Try

            Select Case True
                Case TypeOf p_Object Is Outlook.MailItem

                    myMailItemHandler = MailItemHandlerList.GetByMailItemID(DirectCast(p_Object, Outlook.MailItem))

                    myMailItemHandler.SaveSalutationToReceipients()

                    MailItemHandlerList.Remove(myMailItemHandler)

                Case Else
                    Return
            End Select

        Catch ex As Exception
            Log.Fatal("ItemSendEvent", ex)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ItemSendEvent")
        End Try

    End Sub

End Class
