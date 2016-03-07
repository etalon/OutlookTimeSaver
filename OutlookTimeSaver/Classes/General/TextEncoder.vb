Public Class TextEncoder

    Private Const ESCAPE_CHAR As Char = "\"c
    Private Const UNICODE_CHAR As Char = "u"c

    Private FCodeTable As Dictionary(Of Char, String)
    Private FUnicodeChars As List(Of Char)

    Public Sub New()

        FCodeTable = New Dictionary(Of Char, String)
        FUnicodeChars = New List(Of Char)
        Define(ESCAPE_CHAR)

    End Sub

    Public Sub Define(ByVal AChar As Char)

        Define(AChar, AChar)

    End Sub

    Public Sub Define(ByVal APlaintext As String, ByVal ACode As Char)

        If ACode = UNICODE_CHAR Then Throw New Exception(String.Format("Das Zeichen '{0}' darf nicht verwendet werden.", UNICODE_CHAR))
        FCodeTable(ACode) = APlaintext

    End Sub

    Public Sub DefineUnicode(ByVal APlaintext As Char)

        If Not FUnicodeChars.Contains(APlaintext) Then FUnicodeChars.Add(APlaintext)

    End Sub

    Public Function Encode(ByVal AValue As String) As String

        Dim LIndex As Integer = 0
        Dim LChar As Char

        For Each LEntry As KeyValuePair(Of Char, String) In FCodeTable
            AValue = AValue.Replace(LEntry.Value, ESCAPE_CHAR & LEntry.Key)
        Next

        While LIndex < AValue.Length
            LChar = AValue(LIndex)
            If Char.IsControl(LChar) Then
                AValue = EncodeUnicodeChar(AValue, LChar)
                LIndex += 6
            Else
                LIndex += 1
            End If
        End While

        If AValue.IndexOfAny(FUnicodeChars.ToArray) >= 0 Then
            For Each LChar In FUnicodeChars
                AValue = EncodeUnicodeChar(AValue, LChar)
            Next
        End If

        Return AValue.ToString

    End Function

    Private Function EncodeUnicodeChar(ByVal AValue As String, ByVal AUnicodeChar As Char) As String

        Return AValue.Replace(AUnicodeChar, ESCAPE_CHAR & UNICODE_CHAR & AscW(AUnicodeChar).ToString("X4"))

    End Function

    Public Function Decode(ByVal AValue As String) As String

        Dim LResult As System.Text.StringBuilder = New System.Text.StringBuilder(AValue.Length)
        Dim LEncoded As Boolean = False
        Dim LTemp As String = ""

        Dim LIndex As Integer = 0
        Dim LChar As Char

        While LIndex < AValue.Length
            LChar = AValue(LIndex)
            LIndex += 1
            If LChar = ESCAPE_CHAR Then
                LChar = AValue(LIndex)
                LIndex += 1
                If LChar = "u"c Then
                    LResult.Append(ChrW(Integer.Parse(AValue.Substring(LIndex, 4), Globalization.NumberStyles.HexNumber)))
                    LIndex += 4
                Else
                    If FCodeTable.TryGetValue(LChar, LTemp) Then LResult.Append(LTemp)
                End If
            Else
                LResult.Append(LChar)
            End If
        End While
        Return LResult.ToString

    End Function

End Class