Public Class MailItemHandlerList

    Private Shared m_MailItemHandlers As New List(Of MailItemHandler)

    Public Shared Sub Add(p_MailItem As Outlook.MailItem)

        m_MailItemHandlers.Add(New MailItemHandler(p_MailItem))

    End Sub

    Public Shared Function GetByMailItemID(p_MailItem As Outlook.MailItem) As MailItemHandler

        Return m_MailItemHandlers.First(Function(x) x.UniqueId = p_MailItem.ConversationIndex)

    End Function

    Public Shared Sub Remove(p_MailItemHandler As MailItemHandler)

        m_MailItemHandlers.Remove(p_MailItemHandler)

    End Sub

End Class
