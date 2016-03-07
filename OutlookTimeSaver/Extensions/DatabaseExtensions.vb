Imports System.Runtime.CompilerServices

Public Module DatabaseExtensions

    <Extension()> _
    Public Function IsDBNull(ByVal this As IDataRecord, ByVal name As String) As Boolean

        Return this.IsDBNull(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function GetInteger(ByVal this As IDataRecord, ByVal index As Integer) As Integer

        Return this.GetInt32(index)

    End Function

    <Extension()> _
    Public Function GetInteger(ByVal this As IDataRecord, ByVal name As String) As Integer

        Return this.GetInteger(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function GetInteger(ByVal this As IDataRecord, ByVal name As String, ByVal defaultValue As Integer) As Integer

        Dim index As Integer

        index = this.GetOrdinal(name)
        Return If(this.IsDBNull(index), defaultValue, this.GetInteger(index))

    End Function

    <Extension()> _
    Public Function GetString(ByVal this As IDataRecord, ByVal name As String) As String

        Return this.GetString(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function GetString(ByVal this As IDataRecord, ByVal name As String, ByVal defaultValue As String) As String

        Dim index As Integer

        index = this.GetOrdinal(name)
        Return If(this.IsDBNull(index), defaultValue, this.GetString(index))

    End Function

    <Extension()> _
    Public Function GetDateTime(ByVal this As IDataRecord, ByVal name As String) As Date

        Return this.GetDateTime(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function GetDecimal(ByVal this As IDataRecord, ByVal name As String) As Decimal

        Return this.GetDecimal(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function GetBoolean(ByVal this As IDataRecord, ByVal name As String) As Boolean

        Return this.GetBoolean(this.GetOrdinal(name))

    End Function

    <Extension()> _
    Public Function CommandTextWithReplacedParameters(ByVal sender As IDbCommand) As String
        Dim sb As New System.Text.StringBuilder(sender.CommandText)
        Dim EmptyParameterNames = (From T In sender.Parameters.Cast(Of IDataParameter)() Where String.IsNullOrEmpty(T.ParameterName)).FirstOrDefault

        If EmptyParameterNames IsNot Nothing Then
            Return sender.CommandText
        End If

        For Each p As IDataParameter In sender.Parameters

            If p.Value Is Nothing Then
                Throw New Exception("Es wurde kein Wert für Parameter '" & p.ParameterName & "' übergeben!")
            End If

            Select Case p.DbType
                Case DbType.AnsiString, _
                     DbType.AnsiStringFixedLength, _
                     DbType.Date, DbType.DateTime, _
                     DbType.DateTime2, DbType.Guid, _
                     DbType.String, _
                     DbType.StringFixedLength, _
                     DbType.Time, _
                     DbType.Xml

                    If p.ParameterName(0) = "@" Then
                        sb = sb.Replace(p.ParameterName, "'" & p.Value.ToString.Replace("'", "''") & "'")
                    Else
                        sb = sb.Replace("@" & p.ParameterName, "'" & p.Value.ToString.Replace("'", "''") & "'")
                    End If
                Case Else
                    If p.ParameterName(0) = "@" Then
                        sb = sb.Replace(p.ParameterName, p.Value.ToString)
                    Else
                        sb = sb.Replace("@" & p.ParameterName, p.Value.ToString)
                    End If
            End Select
        Next

        Return sb.ToString

    End Function

End Module
