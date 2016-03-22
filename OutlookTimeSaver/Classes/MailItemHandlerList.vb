Public Class MailItemHandlerList

    Private Shared m_MailItemHandlers As New List(Of MailItemHandler)

    Public Shared Sub Add(p_MailItem As Outlook.MailItem)

        m_MailItemHandlers.Add(New MailItemHandler(p_MailItem))

    End Sub

    Public Shared Sub Remove(p_MailItemHandler As MailItemHandler)

        m_MailItemHandlers.Remove(p_MailItemHandler)

    End Sub

End Class
