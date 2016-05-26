Imports System.Runtime.CompilerServices

Public Module DateExtensions

    ''' <summary>
    ''' Überprüft ob die Uhrzeit zwischen startTime und endTime liegt.
    ''' Es ist auch möglich über Mitternacht zu prüfen, wenn startTime größer als endTime ist.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <param name="startTime"></param>
    ''' <param name="endTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function IsBetween(ByVal this As Date, ByVal startTime As TimeSpan, ByVal endTime As TimeSpan) As Boolean

        Dim currentTime As TimeSpan

        currentTime = this.TimeOfDay
        If startTime < endTime Then
            Return currentTime >= startTime AndAlso currentTime <= endTime
        Else
            Return currentTime >= startTime OrElse currentTime <= endTime
        End If

    End Function

End Module
