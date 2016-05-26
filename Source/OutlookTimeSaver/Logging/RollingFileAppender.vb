Imports System.IO

Module RollingFileAppender

    Public Function MyRollingFileAppender() As log4net.Appender.RollingFileAppender

        Dim fileAppender = New log4net.Appender.RollingFileAppender()

        Dim patternLayout As New log4net.Layout.PatternLayout()
        patternLayout.ConversionPattern = "%date{yyyy-MM-dd HH:mm:ss} - %thread - %-5level %logger - %message%newline"
        patternLayout.ActivateOptions()

        With fileAppender
            .AppendToFile = True
            .Threshold = log4net.Core.Level.Debug
            .File = Path.Combine(Path.Combine(AppDataPath, "Logs"), "Log")
            .DatePattern = "yyyy-MM-dd'.log'"
            .StaticLogFileName = False
            .Layout = patternLayout
            .RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Date
            .ActivateOptions()

        End With

        Return fileAppender

    End Function

End Module
