Imports System.Runtime.InteropServices

Public Class DeletedItemsHandler

    Private m_TrayFolder As Outlook.MAPIFolder
    Private WithEvents m_TrayItems As Outlook.Items

    Public Sub New(p_Application As Outlook.Application)

        Try

            m_TrayFolder = p_Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
            m_TrayItems = m_TrayFolder.Items

        Catch ex As Exception
            Log.Fatal("DeletedItemsHandler.New", ex)
            MessageBox.Show(ex.ToString, "DeletedItemsHandler", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub TrayItems_ItemAdd(p_Item As Object) Handles m_TrayItems.ItemAdd

        Dim mail As Outlook.MailItem

        Log.Debug("TrayItems_ItemAdd - Start")

        If TypeOf p_Item Is Outlook.MailItem Then

            mail = DirectCast(p_Item, Outlook.MailItem)
            With mail
                If .UserProperties.Find("HardDelete") IsNot Nothing Then
                    Log.Debug("Mail vollständig löschen, da es nur ein Entwurf war...")
                    .Delete()
                    Return
                End If

                If Config.My.UnReadDeletedItems Then
                    .UnRead = False
                End If

            End With

            Marshal.ReleaseComObject(mail)

        End If

        Log.Debug("TrayItems_ItemAdd - Ende")

    End Sub

End Class
