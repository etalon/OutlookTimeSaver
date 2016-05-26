Public Class EtiPath

    Public Shared Function IsPathDelimiter(ByVal chr As Char) As Boolean

        Return chr = IO.Path.DirectorySeparatorChar Or chr = IO.Path.AltDirectorySeparatorChar

    End Function

    Public Shared Function IncludeTrailingPathDelimiter(ByVal path As String) As String

        If String.IsNullOrEmpty(path) Then Return String.Empty
        If Not IsPathDelimiter(path.Last) Then path &= IO.Path.DirectorySeparatorChar
        Return path

    End Function

    Public Shared Function ExcludeTrailingPathDelimiter(ByVal path As String) As String

        If String.IsNullOrEmpty(path) Then Return String.Empty
        If IsPathDelimiter(path.Last) Then path = path.Substring(0, path.Length - 1)
        Return path

    End Function

    Public Shared Function RelativeToAbsFile(ByVal path As String, Optional ByVal base As String = Nothing) As String

        If String.IsNullOrEmpty(path) Then Return base
        If String.IsNullOrEmpty(base) Then base = Application.StartupPath
        Return ExcludeTrailingPathDelimiter(IO.Path.GetFullPath(IO.Path.Combine(base, path)))

    End Function

    Public Shared Function RelativeToAbsDir(ByVal path As String, Optional ByVal base As String = Nothing) As String

        Return IncludeTrailingPathDelimiter(RelativeToAbsFile(path, base))

    End Function

    Public Shared Function AbsToRelativeFile(ByVal path As String, Optional ByVal base As String = Nothing) As String

        Return ExcludeTrailingPathDelimiter(AbsToRelativeDir(path, base))

    End Function

    Public Shared Function AbsToRelativeDir(ByVal path As String, Optional ByVal base As String = Nothing) As String

        If String.IsNullOrEmpty(path) Then Return String.Empty
        If Not IO.Path.IsPathRooted(path) Then Return path
        If String.IsNullOrEmpty(base) Then base = Application.StartupPath
        path = IncludeTrailingPathDelimiter(path)
        base = IncludeTrailingPathDelimiter(base)
        path = Uri.UnescapeDataString(New Uri(base).MakeRelativeUri(New Uri(path)).ToString)
        Return path.Replace(IO.Path.AltDirectorySeparatorChar, IO.Path.DirectorySeparatorChar)

    End Function

    Public Shared Function SamePath(ByVal path1 As String, ByVal path2 As String) As Boolean

        path1 = ExcludeTrailingPathDelimiter(IO.Path.GetFullPath(path1))
        path2 = ExcludeTrailingPathDelimiter(IO.Path.GetFullPath(path2))
        Return path1.SameText(path2)

    End Function

End Class