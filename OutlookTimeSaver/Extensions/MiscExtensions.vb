Imports System.Runtime.CompilerServices

Public Module MiscExtensions

    ''' <summary>
    ''' Gibt an ob es sich bei dem Typ um einen skalaren Datentyp handelt.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function IsScalar(ByVal this As Type) As Boolean

        Return this.IsPrimitive OrElse this.Equals(GetType(String))

    End Function

    ''' <summary>
    ''' Gibt bei einem Nullable den darunter liegenden Typ zurück, andernfalls sich selbst.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function ResolveNullableType(ByVal this As Type) As Type

        If this.IsGenericType AndAlso this.GetGenericTypeDefinition.Equals(GetType(Nullable(Of ))) Then
            this = Nullable.GetUnderlyingType(this)
        End If
        Return this

    End Function

End Module
