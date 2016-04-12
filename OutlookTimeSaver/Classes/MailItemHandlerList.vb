﻿Public Class MailItemHandlerList

    Private Shared m_MailItemHandlers As New List(Of MailItemHandler)

    Public Shared Sub Add(p_MailItem As Outlook.MailItem, p_IsInlineResponse As Boolean)

        m_MailItemHandlers.Add(New MailItemHandler(p_MailItem, p_IsInlineResponse))

    End Sub

    Public Shared Sub Remove(p_MailItemHandler As MailItemHandler)

        m_MailItemHandlers.Remove(p_MailItemHandler)

    End Sub

End Class
