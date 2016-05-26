Imports System.Runtime.CompilerServices

Public Module StringExtensions

    ''' <summary>
    ''' Gibt an, ob der übergebene String den gleichen Inhalt hat. Dabei wird die Groß- und Kleinschreibung ignoriert.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SameText(ByVal this As String, ByVal value As String) As Boolean

        Return String.Compare(this, value, True) = 0

    End Function

    <Extension()> _
    Public Function SplitInto(ByVal this As String, ByRef left As String, ByRef right As String, ByVal separator As Char) As Boolean

        Dim index As Integer

        index = this.IndexOf(separator)
        If index < 0 Then
            Return False
        Else
            left = this.Substring(0, index)
            right = this.Substring(index + 1)
            Return True
        End If

    End Function

    ''' <summary>
    ''' Erstellt eine direkte Kopie des angegebenen Arrays.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function GetCopy(ByVal this As String()) As String()

        Return DirectCast(this.Clone, String())

    End Function

End Module
