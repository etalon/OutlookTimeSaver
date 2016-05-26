Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class EnumClass

#Region "Allgemein"

    ''' <summary>
    ''' Ermittelt die Namen aus einer Enum
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <returns></returns>
    ''' <remarks>Beispielaufruf : Dim Zeile() as string = clsEnum.GetNames(GetType(POS_LT11)) </remarks>
    Public Shared Function GetNames(ByVal enumType As Type) As String()
        Return GetNames(GetFieldInfo(enumType))
    End Function

    Friend Shared Function GetNames(ByVal fi As FieldInfo()) As String()

        Dim fieldNames(fi.Length - 1) As String

        For i As Integer = 0 To fi.Length - 1
            fieldNames(i) = fi(i).Name
        Next

        Return fieldNames

    End Function

    ''' <summary>
    ''' Ermittelt die Integer-Werte aus einer Enum
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <returns></returns>
    ''' <remarks>Beispielaufruf : Dim Zeile() as string = clsEnum.GetNames(GetType(POS_LT11)) </remarks>
    Public Shared Function GetIndexes(ByVal enumType As Type) As Integer()
        Return GetIndexes(GetFieldInfo(enumType))
    End Function

    Friend Shared Function GetIndexes(ByVal fi As FieldInfo()) As Integer()

        Dim fieldIndexes(fi.Length - 1) As Integer

        For i As Integer = 0 To fi.Length - 1
            fieldIndexes(i) = CInt(fi(i).GetValue(Nothing))
        Next

        Return fieldIndexes

    End Function

    ''' <summary>
    ''' Prüft ob ein Element enthalten ist
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <param name="enumElement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Contains(ByVal enumType As Type, ByVal enumElement As Integer) As Boolean
        Dim fi As FieldInfo() = getFieldInfo(enumType)
        For i As Integer = 0 To fi.Length - 1
            If CInt(fi(i).GetValue(Nothing)) = enumElement Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Gibt einen Namen aus einer Enum zurück
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <param name="enumElement"></param>
    ''' <returns></returns>
    ''' <remarks>Beispielaufruf : Dim Feld as string = clsEnum.GetNames(GetType(POS_LT11),POS_LT11.VLTYP) </remarks>
    Public Shared Function GetName(ByVal enumType As Type, ByVal enumElement As Integer) As String
        Dim fi As FieldInfo() = getFieldInfo(enumType)
        Return fi(enumElement).Name()
    End Function

    ''' <summary>
    ''' Ermittelt einen Wert aus einer Enum über den Namen
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <param name="szName"></param>
    ''' <returns></returns>
    ''' <remarks>Beispielaufruf : Dim Feld as string = clsEnum.GetNames(GetType(POS_LT11),"VLTYP") </remarks>
    Public Shared Function GetIndex(ByVal enumType As Type, ByVal szName As String) As Integer
        Return GetIndex(getFieldInfo(enumType), szName)
    End Function

    Friend Shared Function GetIndex(ByVal fi As FieldInfo(), ByVal szName As String) As Integer

        For i As Integer = 0 To fi.Length - 1
            If fi(i).Name.ToLower = szName.ToLower Then
                Return CInt(fi(i).GetValue(Nothing))
            End If
        Next

        Return -1

    End Function

    ''' <summary>
    ''' Ermittelt die Anzahl der Enum-Einträge
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Count(ByVal enumType As Type) As Integer
        Return getFieldInfo(enumType).length()
    End Function

#End Region

#Region "Prompt/String-Handling"

    ''' <summary>
    ''' Ersetzen von Platzhaltern über ein Array und zugehörige Enum. Platzhalter müssen mit eckigen Klammern angeführt sein
    ''' </summary>
    ''' <param name="enumType">GetType(Enum)</param>
    ''' <param name="fieldArray">Array, was gemäß Enum aufgebaut ist</param>
    ''' <param name="expression ">String in dem ersetzt werden soll</param>
    ''' <returns></returns>
    ''' <remarks>Speziell zum Füllen z.B. von Prompts mit suchen und ersetzen aller Elemente einer Enum aus einem Array
    ''' Beispielaufruf : szPrompt = clsEnum.SetValues(GetType(POS_LT11), Zeile, szPrompt) 
    ''' Wobei POS_LT11 eine Enum ist, Zeile in String() mit Enum Feldern laut POS_LT11 und 
    ''' szPrompt eine Maske, in die die Felder über den Enum-Namen eingetragen werden sollen.</remarks>
    Public Shared Function ReplacePlaceholders(ByVal enumType As Type, ByVal fieldArray As String(), ByVal expression As String) As String

        Dim fi As FieldInfo() = GetFieldInfo(enumType)
        Dim placeHolderName As String
        Dim idx As Integer

        For Each oMatch As Match In Regex.Matches(expression, "\[\w+\]", RegexOptions.IgnoreCase)

            placeHolderName = oMatch.Value.TrimStart("["c).TrimEnd("]"c)
            idx = EnumClass.GetIndex(fi, placeHolderName)

            If idx < 0 Then
                Continue For
            End If

            expression = Replace(expression, oMatch.Value, fieldArray(idx))

        Next

        Return expression

    End Function

#End Region

#Region "Objekt-Handling"

    ''' <summary>
    ''' Setzt in einem Array über die Namen in enum die Inhalte aus den gleichnamigen Elementen des Objects
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <param name="fieldArray"></param>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SetArrayValues(ByVal enumType As Type, ByRef fieldArray As String(), ByVal obj As Object) As Boolean
        Dim fi As FieldInfo() = GetFieldInfo(enumType)

        For i As Integer = 0 To fi.Length - 1
            Try
                fieldArray(CInt(fi(i).GetValue(Nothing))) = obj.GetType.GetField(fi(i).Name, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase).GetValue(obj).ToString
            Catch ex As Exception
            End Try
        Next
        Return True
    End Function

    ''' <summary>
    ''' Füllt ein übergebenes Object (Class mit Public's) mit den Werten eines Arrays über eine Enum
    ''' </summary>
    ''' <param name="enumType"></param>
    ''' <param name="fieldArray"></param>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SetObjValues(Of T)(ByVal enumType As Type, ByVal fieldArray As String(), ByRef obj As T) As Boolean
        Dim fi As FieldInfo() = GetFieldInfo(enumType)
        Dim fo As FieldInfo() = obj.GetType.GetFields()

        Try
            For Each field As FieldInfo In fo
                Dim szName As String = field.Name

                For n As Integer = 0 To fi.Length - 1
                    If fi(n).Name.ToLower = szName.ToLower Then
                        Dim szHelp As String = fieldArray(CInt(fi(n).GetValue(Nothing)))
                        If Not IsNothing(szHelp) Then
                            field.SetValue(obj, szHelp)
                        End If
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
        End Try

        Return True
    End Function

#End Region

#Region "Mapping-Handling"

    Public Const sMappingRD As String = "~$~"
    Public Const sMappingFD As String = "="

    ''' <summary>
    ''' Wandelt ein String-Array in einen MappingString um. Beispiel:
    ''' MATNR=4711
    ''' CHARG=0815 etc...
    ''' Jeder Wert aus dem Enum wird mit dem entsprechenden Namen versehen
    ''' </summary>
    ''' <param name="enumType">Enum-Typ</param>
    ''' <param name="aBuf">Umzuwandelndes Array</param>
    ''' <returns>MappingString</returns>
    ''' <remarks></remarks>
    Public Shared Function GetMappingStringFromArray(ByVal enumType As Type, ByVal aBuf() As String) As String

        Dim sBuf As String = ""
        Dim fi As FieldInfo()

        Try
            fi = GetFieldInfo(enumType)

            For i As Integer = 0 To fi.Length - 1
                sBuf &= fi(i).Name.ToLower & sMappingFD & aBuf(i) & sMappingRD
            Next

            If sBuf.EndsWith(sMappingRD) Then
                sBuf = sBuf.Substring(0, sBuf.Length - sMappingRD.Length)
            End If

            Return sBuf

        Catch ex As Exception
            Throw New Exception("clsEnum.GetMappingStringFromArray(): " & ex.Message)
        End Try

    End Function

    ''' <summary>
    ''' Versucht einen MappingString in ein String-Array umzuwandeln. Es muss aber nicht jedes Feld vorhanden sein
    ''' </summary>
    ''' <param name="enumType">Enum-Typ</param>
    ''' <param name="sMappingString">MappingString</param>
    ''' <returns>String-Array</returns>
    ''' <remarks></remarks>
    Public Shared Function GetArrayFromMappingString(ByVal enumType As Type, ByVal sMappingString As String) As String()

        Dim fi As FieldInfo()
        Dim aBuf() As String
        Dim aEntry() As String

        Try
            fi = GetFieldInfo(enumType)

            ReDim aBuf(fi.Length - 1)
            For i As Integer = 0 To aBuf.Length - 1
                aBuf(i) = ""
            Next

            For Each sEntry As String In Split(sMappingString, sMappingRD)
                aEntry = Split(sEntry, sMappingFD, 2)

                Try
                    aBuf(GetIndex(enumType, aEntry(0))) = aEntry(1)
                Catch ex As Exception
                    ' Wenn kein Index ermittelt werden konnte, haben wir diesen Namen in der aktuellen Enum nicht
                End Try

            Next

            Return aBuf

        Catch ex As Exception
            Throw New Exception("clsEnum.GetArrayFromMappingString(): " & ex.Message, ex)
        End Try

    End Function

#End Region

#Region "Helper"

    Friend Shared Function GetFieldInfo(ByVal enumType As Type) As FieldInfo()
        Return enumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
    End Function

#End Region

#Region "Obsolete-Support"

    ''' <summary>
    ''' Ersetzen von Platzhaltern über ein Array und zugehörige Enum. Platzhalter müssen mit eckigen Klammern angeführt sein
    ''' </summary>
    ''' <param name="enumType">GetType( Enum )</param>
    ''' <param name="fieldArray">Array, was gemäß Enum aufgebaut ist</param>
    ''' <param name="expression ">String in dem ersetzt werden soll</param>
    ''' <returns></returns>
    ''' <remarks>Speziell zum Füllen z.B. von Prompts mit suchen und ersetzen aller Elemente einer Enum aus einem Array
    ''' Beispielaufruf : szPrompt = clsEnum.SetValues(GetType(POS_LT11), Zeile, szPrompt) 
    ''' Wobei POS_LT11 eine Enum ist, Zeile in String() mit Enum Feldern laut POS_LT11 und 
    ''' szPrompt eine Maske, in die die Felder über den Enum-Namen eingetragen werden sollen.</remarks>
    <Obsolete("Bitte ReplacePlaceholders verwenden; 19.12.2013; JH")> _
    Public Shared Function SetValues(ByVal enumType As Type, ByVal fieldArray As String(), ByVal expression As String) As String
        Return ReplacePlaceholders(enumType, fieldArray, expression)
    End Function

#End Region

End Class

''' <summary>
''' Diese Klasse ist eine Art Kapselung, sodass in Modulen ein kürzerer Aufruf bei mehrfacher Wiederholung erfolgen kann
''' Erspart eigentlich nur Tipparbeit
''' </summary>
''' <remarks></remarks>
Public Class clsLoadedEnum

    Private ReadOnly m_EnumType As Type
    Private ReadOnly m_FieldName() As String
    Private ReadOnly m_FieldIndex() As Integer
    Private m_EmptyArray() As String

    Public Sub New(ByVal p_EnumType As Type)

        Dim fi() As FieldInfo

        m_EnumType = p_EnumType

        fi = EnumClass.GetFieldInfo(m_EnumType)

        m_FieldName = EnumClass.GetNames(fi)
        m_FieldIndex = EnumClass.GetIndexes(fi)

    End Sub

    Public ReadOnly Property Type() As Type
        Get
            Return m_EnumType
        End Get
    End Property

    ''' <summary>
    ''' Gibt den Namen zu einem Enum-Eintrag (Index) zurück
    ''' </summary>
    ''' <param name="enumElement">Index</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name(ByVal enumElement As Integer) As String
        Get
            Return m_FieldName(enumElement)
        End Get
    End Property

    ''' <summary>
    ''' Gibt alle Namen einer Enum zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Names() As String()
        Get
            ' Hier ein Clone, weil man sonst von außen das Array ändern kann!
            Return DirectCast(m_FieldName.Clone, String())
        End Get
    End Property

    ''' <summary>
    ''' Ermittelt den Index zu einem Namen aus der Enum
    ''' </summary>
    ''' <param name="pFieldName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Index(ByVal pFieldName As String) As Integer
        Get
            For i As Integer = 0 To Count - 1
                If m_FieldName(i).ToLower = pFieldName.ToLower Then
                    Return i
                End If
            Next

            Return -1

        End Get
    End Property

    Public ReadOnly Property Contains(ByVal pIndex As Integer) As Boolean
        Get
            If Array.IndexOf(m_FieldIndex, pIndex, 0) >= 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>
    ''' Anzahl an Enum-Elementen...
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer
        Get
            Return m_FieldName.Length
        End Get
    End Property

    ''' <summary>
    ''' Liefert ein zur Enum passendes Array mit zurück, wo die Inhalte schon auf Leerstring gesetzt sind
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetEmptyArray() As String()

        If IsNothing(m_EmptyArray) Then
            ReDim m_EmptyArray(Count - 1)
            For i As Integer = 0 To m_EmptyArray.Length - 1
                m_EmptyArray(i) = ""
            Next
        End If

        Return DirectCast(m_EmptyArray.Clone, String())

    End Function

    Public Delegate Function DoTranslationDelegate(ByVal textToTranslate As String) As String

    ''' <summary>
    ''' Gibt einen String zurück, der zu allen Namen in der Enum, den jeweiligen Wert aus dem Array schreibt
    ''' </summary>
    ''' <param name="p_Arr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDebugString(ByVal p_Arr() As String, Optional ByVal p_DoTranslationDelegate As DoTranslationDelegate = Nothing) As String

        Dim sb As New System.Text.StringBuilder(20)

        For i As Integer = 0 To m_FieldName.Length - 1
            If p_DoTranslationDelegate Is Nothing Then
                sb.AppendLine(m_FieldName(i) & " = '" & p_Arr(i) & "'")
            Else
                sb.AppendLine(p_DoTranslationDelegate(m_FieldName(i)) & " = '" & p_Arr(i) & "'")
            End If
        Next

        Return sb.ToString

    End Function

    Public Function ReplacePlaceholders(ByVal fieldArray() As String, ByVal expression As String) As String

        Return EnumClass.ReplacePlaceholders(m_EnumType, fieldArray, expression)

    End Function

End Class