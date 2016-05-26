Public Class MailItemHandlerList

    Private Shared m_MailItemHandlers As New List(Of MailItemHandler)

    Public Shared Sub Add(p_MailItem As Outlook.MailItem, p_IsInlineResponse As Boolean, p_OpenedFromDrafts As Boolean)

        m_MailItemHandlers.Add(New MailItemHandler(p_MailItem, p_IsInlineResponse, p_OpenedFromDrafts))

    End Sub

    Public Shared Sub Remove(p_MailItemHandler As MailItemHandler)

        m_MailItemHandlers.Remove(p_MailItemHandler)

    End Sub

    Public Shared Function Exists(p_EntryId As String) As Boolean

        Return m_MailItemHandlers.Any(Function(x) x.EntryId = p_EntryId)

    End Function

    Public Shared Sub Remove(p_EntryId As String)

        Dim mailItemHandlerObj As MailItemHandler

        mailItemHandlerObj = m_MailItemHandlers.FirstOrDefault(Function(x) x.EntryId = p_EntryId)

        If mailItemHandlerObj Is Nothing Then
            Return
        Else
            m_MailItemHandlers.Remove(mailItemHandlerObj)
        End If

    End Sub

    Public Shared Function GetItem(p_EntryId As String) As MailItemHandler

        Return m_MailItemHandlers.First(Function(x) x.EntryId = p_EntryId)

    End Function

    Public Shared Function TryGetItem(p_EntryId As String, ByRef p_MailItemHandler As MailItemHandler) As Boolean

        p_MailItemHandler = m_MailItemHandlers.FirstOrDefault(Function(x) x.EntryId = p_EntryId)
        Return p_MailItemHandler IsNot Nothing

    End Function

End Class
