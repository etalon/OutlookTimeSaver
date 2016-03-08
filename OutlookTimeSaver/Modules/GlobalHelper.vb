Imports System.Environment
Imports System.IO

Module GlobalHelper

    Public VALID_SALUTATIONS As String() = {"hi", "hallo", "guten morgen", "guten tag", "guten abend", "sehr geehrter", "sehr geehrte"}

    Public Log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public ReadOnly Property ApplicationStartupPath() As String
        Get
            Return System.IO.Path.GetDirectoryName(Application.ExecutablePath)
        End Get
    End Property

    Public ReadOnly Property AppDataPath As String
        Get
            Static myAppDataPath As String = ""

            If String.IsNullOrEmpty(myAppDataPath) Then

                myAppDataPath = Path.Combine(GetFolderPath(SpecialFolder.ApplicationData), "OutlookTimeSaver")
                If Not Directory.Exists(myAppDataPath) Then
                    Directory.CreateDirectory(myAppDataPath)
                End If

            End If

            Return myAppDataPath

        End Get
    End Property

    Public Function ReadTextFileFromResources(p_FileName As String) As String

        Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()

        Using stream As Stream = assembly.GetManifestResourceStream(p_FileName)
            Using reader As New StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using

    End Function

End Module
