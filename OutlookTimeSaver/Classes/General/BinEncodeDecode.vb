
Public Class BinEncodeDecode

    Private Const FD As String = "§"
    Private Const RD As String = "@"

    Public Shared Function EncodeField(ByVal szString As String) As String
        Return BinEncode(szString)
    End Function

    Public Shared Function Encode(ByVal szString As String) As String
        Dim Lines() As String
        Dim Line() As String
        Dim nLines As Integer
        Dim nLine As Integer
        Dim bMultilines As Boolean
        Dim bMultiline As Boolean
        Dim szRet As String
        If InStr(szString, RD) > 0 Then bMultilines = True
        Lines = Split(szString, RD)
        For nLines = 0 To UBound(Lines)
            If InStr(Lines(nLines), FD) > 0 Then bMultiline = True
            Line = Split(Lines(nLines), FD)
            For nLine = 0 To UBound(Line)
                Line(nLine) = BinEncode(Line(nLine))
            Next
            If bMultiline = True Then
                Lines(nLines) = Join(Line, FD)
            Else
                Lines(nLines) = Line(0)
            End If
        Next
        If bMultilines = True Then
            szRet = Join(Lines, RD)
        Else
            szRet = Lines(0)
        End If
        Return szRet
    End Function

    ''' <summary>
    ''' Funktion, welche ein Passwort entschlüsselt. Wenn es nicht verschlüsselt ist, wird dieses verschlüsselt.
    ''' Zur Erkennung ob es verschlüsselt wurde, dient der Parameter "encryptionDone". Wenn True, kann man die Konfigdatei entsprechend aktualisieren
    ''' </summary>
    ''' <param name="cryptedPassword"></param>
    ''' <param name="encryptionDone"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetDecryptedPassword(ByRef cryptedPassword As String, Optional ByRef encryptionDone As Boolean = False) As String

        Const CRYPT_SIGN As String = "[CRYPTED]"

        Dim myPassword As String = cryptedPassword

        encryptionDone = False

        If myPassword.StartsWith(CRYPT_SIGN) Then
            myPassword = Replace(myPassword, CRYPT_SIGN, "", , 1)
            Return BinEncodeDecode.Decode(myPassword)
        ElseIf String.IsNullOrEmpty(myPassword) Then
            Return myPassword
        Else
            encryptionDone = True
            cryptedPassword = CRYPT_SIGN & BinEncodeDecode.Encode(myPassword)
            Return myPassword
        End If

    End Function


    Public Shared Function DecodeField(ByVal szString As String) As String
        Return BinDecode(szString)
    End Function

    Public Shared Function Decode(ByVal szString As String) As String
        Dim Lines() As String
        Dim Line() As String
        Dim nLines As Integer
        Dim nLine As Integer
        Dim bMultilines As Boolean
        Dim bMultiline As Boolean
        Dim szRet As String
        If InStr(szString, RD) > 0 Then bMultilines = True
        Lines = Split(szString, RD)
        For nLines = 0 To UBound(Lines)
            If InStr(Lines(nLines), FD) > 0 Then bMultiline = True
            Line = Split(Lines(nLines), FD)
            For nLine = 0 To UBound(Line)
                Line(nLine) = BinDecode(Line(nLine))
            Next
            If bMultiline = True Then
                Lines(nLines) = Join(Line, FD)
            Else
                Lines(nLines) = Line(0)
            End If
        Next
        If bMultilines = True Then
            szRet = Join(Lines, RD)
        Else
            szRet = Lines(0)
        End If
        Return szRet
    End Function

    Private Shared Function BinDecode(ByVal szString As String) As String
        Dim szHelp As String = ""
        Dim nCount As Integer = 0
        Dim szRet As String = ""

        'Reihenfolge der HEX-Buchstaben tauschen (gemein, was...)
        szHelp = ""
        For nCount = 2 To Len(szString) + 1 Step 2
            szHelp += Mid(szString, nCount, 1) + Mid(szString, nCount - 1, 1)
        Next

        szString = szHelp

        szHelp = ""
        For nCount = 0 To Len(szString) - 1 Step 2
            szHelp = "&H" & szString.Substring(nCount, 2)
            szRet += Chr(CInt(szHelp) Xor 173)
        Next

        'Reihenfolge der Buchstaben tauschen
        szHelp = ""
        For nCount = 2 To Len(szRet) + 1 Step 2
            szHelp += Mid(szRet, nCount, 1) + Mid(szRet, nCount - 1, 1)
        Next

        Return szHelp

    End Function

    Private Shared Function BinEncode(ByVal szString As String) As String
        Dim szHelp As String = ""
        Dim nCount As Integer = 0
        Dim szRet As String = ""

        'Reihenfolge der Buchstaben tauschen
        szHelp = ""
        For nCount = 2 To Len(szString) + 1 Step 2
            szHelp += Mid(szString, nCount, 1) + Mid(szString, nCount - 1, 1)
        Next

        szString = szHelp

        szHelp = ""

        For nCount = 0 To Len(szString) - 1
            szHelp = Right("00" + Hex$(Asc(szString.Substring(nCount, 1)) Xor 173), 2)
            szRet += szHelp
        Next

        'Reihenfolge der HEX-Buchstaben tauschen (gemein, was...)
        szHelp = ""
        For nCount = 2 To Len(szRet) + 1 Step 2
            szHelp += Mid(szRet, nCount, 1) + Mid(szRet, nCount - 1, 1)
        Next

        Return szHelp

    End Function

End Class
