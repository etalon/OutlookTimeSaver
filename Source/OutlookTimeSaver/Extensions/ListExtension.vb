Imports System.Runtime.CompilerServices

Public Module ListExtension

    ''' <summary>
    ''' Verschiebt einen Eintrag an das Ende der Liste
    ''' </summary>
    ''' <param name="this"></param>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    <Extension()> _
    Public Sub MoveToEnd(ByVal this As IList, ByVal obj As Object)

        this.Remove(obj)
        this.Add(obj)

    End Sub

    ''' <summary>
    ''' Konvertiert ein Dictionary in eine Liste die lediglich die Values, jedoch nicht die Keys, enthält.
    ''' </summary>
    ''' <typeparam name="TKey"></typeparam>
    ''' <typeparam name="TValue"></typeparam>
    ''' <param name="this"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function ToValueList(Of TKey, TValue)(ByVal this As IDictionary(Of TKey, TValue)) As List(Of TValue)

        Dim result As New List(Of TValue)

        For Each item As KeyValuePair(Of TKey, TValue) In this
            result.Add(item.Value)
        Next
        Return result

    End Function

    <Extension()> _
    Public Function IsEmpty(ByVal this As IList) As Boolean

        Return this.Count = 0

    End Function

    <Extension()> _
    Public Function Add(Of T)(ByVal this As List(Of T)) As T

        this.Add(Activator.CreateInstance(Of T)())
        Return this.Last

    End Function

End Module
