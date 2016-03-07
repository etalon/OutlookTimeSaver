Imports log4net.Appender
Imports log4net.Core

Public Class LoggerFormAppender
    Implements log4net.Appender.IAppender

    Private m_Name As String = "LoggerFormAppender"
    Private m_Form As LoggerForm

    Public Shared LoggingEvents As New List(Of LoggingEvent)

    Public Property Name As String Implements IAppender.Name
        Get
            Return m_Name
        End Get
        Set(value As String)
            m_Name = value
        End Set
    End Property

    Public Sub New()

        m_Form = New LoggerForm
        With m_Form
            .Show()
            .SetLayoutPosition()
        End With

    End Sub

    Public Sub Close() Implements IAppender.Close

        m_Form.Close()
        m_Form.Dispose()

    End Sub

    Public Sub DoAppend(loggingEvent As LoggingEvent) Implements IAppender.DoAppend

        LoggingEvents.Insert(0, loggingEvent)
        m_Form.RefreshData()

    End Sub
End Class
