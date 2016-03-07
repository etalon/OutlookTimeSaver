Public Class OutlookContacts

    Private Shared m_Contacts As New Dictionary(Of String, Outlook.ContactItem)
    Private Shared m_ContactsFolder As Outlook.MAPIFolder
    Private Shared m_SuggestedContactsFolder As Outlook.MAPIFolder

    Public Shared Sub ReadContacts(p_Application As Outlook.Application)

        m_Contacts.Clear()

        m_ContactsFolder = p_Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        importContactsFromFolder(m_ContactsFolder)

    End Sub

    Private Shared Sub importContactsFromFolder(p_ContactFolder As Outlook.MAPIFolder)

        ' TODO: So viele Items darf man nicht zwischenpuffern
        ' Und außerdem verzögert das den Start von OUtlook.
        ' Daher sollte dies in einem eigenen Thread laufen und man sollte
        ' eine Datenbank aufbauen, welche abgeglichen wird.

        'For Each item As Outlook.ContactItem In p_ContactFolder.Items

        '    If String.IsNullOrEmpty(item.Email1Address) Then
        '        Continue For
        '    End If

        '    m_Contacts.Add(item.Email1Address, item)

        'Next

    End Sub

    Public Shared Function TryGetContact(p_Email As String, ByRef p_Contact As Outlook.ContactItem) As Boolean
        Return m_Contacts.TryGetValue(p_Email, p_Contact)
    End Function

End Class
