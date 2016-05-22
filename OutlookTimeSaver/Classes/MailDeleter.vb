Imports System.Runtime.InteropServices

Public Class MailDeleter

    Private Shared m_MailsToDelete As New List(Of MailToDelete)
    Private Shared m_MailsToDeleteSyncLock As New Object
    Private Shared m_DraftsFolder As Outlook.MAPIFolder

    Public Shared Sub Init(p_Application As Outlook.Application)

        m_DraftsFolder = p_Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)

    End Sub

    Public Shared Sub Add(p_MailItem As Outlook.MailItem)

        SyncLock m_MailsToDeleteSyncLock
            m_MailsToDelete.Add(New MailToDelete(p_MailItem))
        End SyncLock

        With New Threading.Thread(AddressOf CheckMailToDelete)
            .Start()
        End With

    End Sub

    Private Shared Sub CheckMailToDelete()

        Try

            System.Threading.Thread.Sleep(15000)

            SyncLock m_MailsToDeleteSyncLock
                For Each m In m_MailsToDelete

                    If Not MailItemHandlerList.GetItem(m.MailItem.EntryID).HasManuallyChanged Then
                        Log.Debug("Mail wurde nicht verändert, also löschen")
                        m.MailItem.UserProperties.Add("HardDelete", Microsoft.Office.Interop.Outlook.OlUserPropertyType.olText)
                        m.MailItem.Save()
                        m.MailItem.Delete()

                        Marshal.ReleaseComObject(m.MailItem)

                    End If

                Next

                m_MailsToDelete.Clear()

            End SyncLock

        Catch ex As Exception
            Log.Error(ex)
        End Try

    End Sub

End Class

Public Class MailToDelete

    Public Property MailItem As Outlook.MailItem
    Public Property LastModificationDate As Date

    Public Sub New(ByRef p_MailItem As Outlook.MailItem)

        MailItem = p_MailItem

    End Sub

End Class

