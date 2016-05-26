Public Class FetchedDataRow

    Private fields As Dictionary(Of String, Object)

    Public Sub New(ByVal dataSource As IDataRecord)

        fields = New Dictionary(Of String, Object)(dataSource.FieldCount, StringComparer.CurrentCultureIgnoreCase)
        For i = 0 To dataSource.FieldCount - 1
            fields.Add(dataSource.GetName(i), If(dataSource.IsDBNull(i), Nothing, dataSource.GetValue(i)))
        Next

    End Sub

    Public Function IsNull(ByVal name As String) As Boolean

        Return fields(name) Is Nothing

    End Function

    Public Function GetValue(Of T)(ByVal name As String) As T

        Dim result As Object = Nothing

        If Not fields.TryGetValue(name, result) Then
            Throw New Exception(String.Format("Spalte '{0}' existiert nicht in der aktuellen Abfrage.", name))
        End If

        Return CType(result, T)

    End Function

    Public Sub GetValue(Of T)(ByVal name As String, ByRef result As T)

        result = GetValue(Of T)(name)

    End Sub

    Public Function GetValueDefault(Of T)(ByVal name As String, ByVal defaultValue As T) As T

        If IsNull(name) Then
            Return defaultValue
        Else
            Return GetValue(Of T)(name)
        End If

    End Function

    Public Sub GetValueDefault(Of T)(ByVal name As String, ByRef result As T, ByVal defaultValue As T)

        result = GetValueDefault(Of T)(name, defaultValue)

    End Sub

End Class