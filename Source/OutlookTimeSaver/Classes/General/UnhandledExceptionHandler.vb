''' <summary>
''' Mit dieser Klasse sollen unbehandelte Exceptions abgefangen werden können.
''' Tritt eine solche Exception auf, wird diese in der ERROR.log vermerkt.
''' </summary>
''' <remarks></remarks>
Public Class UnhandledExceptionHandler

    Public Shared Sub Activate()

        AddHandler Application.ThreadException, AddressOf ThreadExceptionHandler
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionHandler

    End Sub

    Private Shared Sub ThreadExceptionHandler(ByVal sender As System.Object, ByVal e As System.Threading.ThreadExceptionEventArgs)

        HandleException(e.Exception, False)

    End Sub

    Private Shared Sub UnhandledExceptionHandler(ByVal sender As System.Object, ByVal e As System.UnhandledExceptionEventArgs)

        HandleException(CType(e.ExceptionObject, Exception), e.IsTerminating)

    End Sub

    Private Shared Sub HandleException(ByVal p_Exception As Exception, ByVal p_Kill As Boolean)

        Try
            Log.Error("Schwerer Fehler", p_Exception)
        Catch
        End Try
        If p_Kill Then System.Environment.Exit(-1)

    End Sub

End Class
