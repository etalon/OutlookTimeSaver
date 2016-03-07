Imports System.IO

Public Class Config

    Private Shared m_ConfigFile As String = Path.Combine(AppDataPath, "OutlookTimeSaverConfig.json")
    Private Shared m_Config As OutlookTimeSaverConfig

    Public Shared ReadOnly Property My As OutlookTimeSaverConfig
        Get
            Return m_Config
        End Get
    End Property

    Public Shared Sub Load()

        If Not File.Exists(m_ConfigFile) Then
            m_Config = New OutlookTimeSaverConfig

            Using sw As New StreamWriter(m_ConfigFile, False, System.Text.Encoding.Default)
                sw.Write(JSONSerializer.DirectSerialize(m_Config, True))
            End Using

        End If

        Using sr As New StreamReader(m_ConfigFile, System.Text.Encoding.Default)
            m_Config = JSONSerializer.DirectDeserialize(Of OutlookTimeSaverConfig)(sr.ReadToEnd)
        End Using

    End Sub

    Public Class OutlookTimeSaverConfig

        Public DebugViewMode As Boolean = True
        Public LoggingEnabled As Boolean = True

    End Class

End Class
