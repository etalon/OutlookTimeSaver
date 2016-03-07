Imports System.Reflection

Public Class JSONSerializer

#Region " Allgemein "

    Public Const NULL As String = "null"

    Protected Enum ReadOption
        Normal
        DontTrim
        Hold
    End Enum

    Public Enum ObjectMembers
        Both
        Fields
        Properties
    End Enum

    Protected Shared FCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.GetCultureInfo("en-US")
    Protected Shared FDateFormat As String = "s" ' ISO 8601

    Protected FHumanReadable As Boolean
    Protected FStrict As Boolean
    Protected FSerializedObjectMembers As ObjectMembers
    Protected FEncoder As TextEncoder
    Protected FTypeBind As Dictionary(Of Type, Type)
    Protected FOutput As StringBuilder
    Protected FLevel As Integer
    Protected FInput As String
    Protected FInputIndex As Integer
    Protected FComment As Boolean

    Public ReadOnly Property Encoder() As TextEncoder
        Get
            Return FEncoder
        End Get
    End Property

    Public ReadOnly Property TypeBind() As Dictionary(Of Type, Type)
        Get
            Return FTypeBind
        End Get
    End Property

    Public Property HumanReadable() As Boolean
        Get
            Return FHumanReadable
        End Get
        Set(ByVal value As Boolean)
            FHumanReadable = value
        End Set
    End Property

    Public Property Strict() As Boolean
        Get
            Return FStrict
        End Get
        Set(ByVal value As Boolean)
            FStrict = value
        End Set
    End Property

    Public Property SerializedObjectMembers() As ObjectMembers
        Get
            Return FSerializedObjectMembers
        End Get
        Set(ByVal value As ObjectMembers)
            FSerializedObjectMembers = value
        End Set
    End Property

    Public Sub New(Optional ByVal AHumanReadable As Boolean = False)

        FHumanReadable = AHumanReadable
        FStrict = True
        FSerializedObjectMembers = ObjectMembers.Both
        FEncoder = New TextEncoder()
        With FEncoder
            .Define(""""c)
            .Define("/"c)
            .Define(vbBack, "b"c)
            .Define(vbFormFeed, "f"c)
            .Define(vbNewLine, "n"c)
            .Define(vbCr, "r"c)
            .Define(vbTab, "t"c)
        End With
        FTypeBind = New Dictionary(Of Type, Type)

    End Sub

#End Region

#Region " Serialisierung "

    Protected Function IsObjectMemberSerializable(ByVal AMember As ObjectMembers) As Boolean

        Return FSerializedObjectMembers = AMember OrElse FSerializedObjectMembers = ObjectMembers.Both

    End Function

    Protected Function IsNumericType(ByVal AValue As Object) As Boolean

        Return (TypeOf AValue Is Decimal) OrElse (TypeOf AValue Is Double) OrElse (TypeOf AValue Is Single) OrElse _
            (TypeOf AValue Is Long) OrElse (TypeOf AValue Is Integer) OrElse (TypeOf AValue Is Short) OrElse (TypeOf AValue Is SByte) OrElse _
            (TypeOf AValue Is ULong) OrElse (TypeOf AValue Is UInteger) OrElse (TypeOf AValue Is UShort) OrElse (TypeOf AValue Is Byte)

    End Function

    Protected Function IsStringType(ByVal AValue As Object) As Boolean

        Return (TypeOf AValue Is String) OrElse (TypeOf AValue Is Char) OrElse (TypeOf AValue Is Version) OrElse _
            (TypeOf AValue Is Type) OrElse (TypeOf AValue Is TimeSpan)

    End Function

    Protected Function EncodeString(ByVal APlaintext As String, Optional ByVal AStrict As Boolean = True) As String

        If APlaintext.All(Function(LChar) Char.IsLetterOrDigit(LChar)) Then
            Return If(AStrict, """"c & APlaintext & """"c, APlaintext)
        Else
            Return """"c & FEncoder.Encode(APlaintext) & """"c
        End If

    End Function

    Protected Sub WriteIndent()

        If FHumanReadable AndAlso FOutput.Length > 0 Then FOutput.Append(vbNewLine & New String(CChar(vbTab), FLevel))

    End Sub

    Protected Sub WriteProperty(ByVal AName As String, ByVal AValue As String)

        WriteIndent()
        If AName IsNot Nothing Then
            FOutput.Append(EncodeString(AName, FStrict) & ":"c)
            If FHumanReadable Then FOutput.Append(" "c)
        End If
        FOutput.Append(AValue)

    End Sub

    Protected Sub WriteArray(ByVal AArray As IEnumerable)

        Dim LFirst As Boolean = True

        FOutput.Append("["c)
        FLevel += 1

        For Each LItem As Object In AArray
            WriteValue(Nothing, LItem, LFirst)
        Next

        FLevel -= 1
        WriteIndent()
        FOutput.Append("]"c)

    End Sub

    Protected Sub WriteObject(ByVal AObject As Object)

        Dim LFirst As Boolean = True
        Dim LDictionary As IDictionary

        FOutput.Append("{"c)
        FLevel += 1

        LDictionary = TryCast(AObject, IDictionary)
        If LDictionary IsNot Nothing Then
            For Each LEntry As DictionaryEntry In LDictionary
                WriteValue(LEntry.Key.ToString, LEntry.Value, LFirst)
            Next
        Else
            If IsObjectMemberSerializable(ObjectMembers.Fields) Then
                For Each LField As FieldInfo In AObject.GetType.GetFields()
                    WriteValue(LField.Name, LField.GetValue(AObject), LFirst)
                Next
            End If
            If IsObjectMemberSerializable(ObjectMembers.Properties) Then
                For Each LProperty As PropertyInfo In AObject.GetType.GetProperties
                    If LProperty.CanRead AndAlso LProperty.GetIndexParameters.Length = 0 Then
                        WriteValue(LProperty.Name, LProperty.GetValue(AObject, Nothing), LFirst)
                    End If
                Next
            End If
        End If

        FLevel -= 1
        WriteIndent()
        FOutput.Append("}"c)

    End Sub

    Protected Sub WriteValue(ByVal AName As String, ByVal AValue As Object, ByRef AFirst As Boolean)

        Dim LType As Type

        If AFirst Then
            AFirst = False
        Else
            FOutput.Append(","c)
        End If

        If AValue Is Nothing Then
            WriteProperty(AName, NULL)
        Else
            LType = AValue.GetType
            Select Case True
                Case TypeOf AValue Is Boolean
                    WriteProperty(AName, AValue.ToString.ToLower)

                Case IsNumericType(AValue)
                    WriteProperty(AName, CDec(AValue).ToString(FCulture))

                Case IsStringType(AValue) OrElse LType.IsEnum
                    WriteProperty(AName, EncodeString(AValue.ToString))

                Case TypeOf AValue Is Date
                    WriteProperty(AName, EncodeString(CDate(AValue).ToString(FDateFormat)))

                Case Not LType.IsPrimitive
                    WriteProperty(AName, "")
                    If LType.IsArray Or TryCast(AValue, IList) IsNot Nothing Then
                        WriteArray(CType(AValue, IEnumerable))
                    Else
                        WriteObject(AValue)
                    End If

                Case Else
                    Throw New Exception(String.Format("Der Datentyp '{0}' wird nicht unterstützt.", LType.Name))

            End Select
        End If

    End Sub

    Public Function Serialize(ByVal AInput As Object) As String

        FLevel = 0
        FOutput = New StringBuilder()
        WriteValue(Nothing, AInput, True)
        Return FOutput.ToString

    End Function

    Public Shared Function DirectSerialize(ByVal AInput As Object, Optional ByVal AHumanReadable As Boolean = False) As String

        With New JSONSerializer(AHumanReadable)
            Return .Serialize(AInput)
        End With

    End Function

#End Region

#Region " Deserialisierung "

    Protected Sub UnexpectedCharFound(ByVal AChar As Char)

        Throw New DeserializationException(Me, "An Position {0} wurde das unerwartete Zeichen '{1}' gefunden.", FInputIndex.ToString, AChar)

    End Sub

    Protected Sub UnexpectedCharFound()

        UnexpectedCharFound(ReadNext)

    End Sub

    Protected Function ReadNext(Optional ByVal AOption As ReadOption = ReadOption.Normal) As Char

        Dim LResult As Char

        Do
            If FInputIndex >= FInput.Length Then Throw New DeserializationException(Me, "Unerwartetes Ende.")
            LResult = FInput(FInputIndex)
            FInputIndex += 1
            If AOption = ReadOption.DontTrim Then Exit Do
            If Not FStrict Then
                If FComment Then
                    If LResult = vbCr OrElse LResult = vbLf Then FComment = False
                Else
                    If LResult = "#"c Then FComment = True
                End If
            End If
        Loop While FComment OrElse Char.IsWhiteSpace(LResult)
        If AOption = ReadOption.Hold Then FInputIndex -= 1
        Return LResult

    End Function

    Protected Function ReadWhenExpected(ByVal AExpectedChar As Char) As Boolean

        If ReadNext(ReadOption.Hold) = AExpectedChar Then
            ReadNext()
            Return True
        Else
            Return False
        End If

    End Function

    Protected Sub CheckNext(ByVal AExpected As Char)

        Dim LChar As Char

        LChar = ReadNext()
        If LChar <> AExpected Then UnexpectedCharFound(LChar)

    End Sub

    Protected Function ReadString() As String

        Dim LChar As Char
        Dim LStartIndex As Integer

        CheckNext(""""c)
        LStartIndex = FInputIndex
        Do
            LChar = ReadNext(ReadOption.DontTrim)
            If LChar = "\"c Then ReadNext(ReadOption.DontTrim)
        Loop Until LChar = """"c
        Return FEncoder.Decode(FInput.Substring(LStartIndex, FInputIndex - LStartIndex - 1))

    End Function

    Protected Function ReadNonString() As String

        Dim LChar As Char
        Dim LStartIndex As Integer

        If FLevel = 0 Then
            ' Wird diese Funktion auf Level 0 aufgerufen, dann wurde nur ein einzelner skalarer Wert zur Deserialisierung übergeben.
            Return FInput.Trim
        Else
            LStartIndex = FInputIndex
            Do
                LChar = ReadNext(ReadOption.DontTrim)
            Loop Until (LChar = ":"c) OrElse (LChar = ","c) OrElse (LChar = "]"c) OrElse (LChar = "}"c) OrElse Char.IsWhiteSpace(LChar)
            FInputIndex -= 1
            Return FInput.Substring(LStartIndex, FInputIndex - LStartIndex)
        End If

    End Function

    Protected Function ReadPropertyName() As String

        Dim LResult As String

        If FStrict OrElse ReadNext(ReadOption.Hold) = """"c Then
            LResult = ReadString()
        Else
            LResult = ReadNonString()
        End If
        CheckNext(":"c)
        Return LResult

    End Function

    Protected Function DecodeStringValue(ByVal AType As Type, ByVal AValue As String) As Object

        Select Case True
            Case GetType(Char).Equals(AType) : Return AValue.Single
            Case GetType(Date).Equals(AType) : Return Date.ParseExact(AValue, FDateFormat, Nothing)
            Case GetType(Version).Equals(AType) : Return New Version(AValue)
            Case GetType(Type).Equals(AType) : Return Type.GetType(AValue, True, True)
            Case GetType(TimeSpan).Equals(AType) : Return TimeSpan.Parse(AValue)
            Case AType.IsEnum : Return [Enum].Parse(AType, AValue, True)
            Case Else : Return AValue
        End Select

    End Function

    Protected Function GenerateArray(ByVal AType As Type, ByVal AValues As IList) As Array

        Dim LResult As Array

        LResult = Array.CreateInstance(AType, AValues.Count)
        AValues.CopyTo(LResult, 0)
        Return LResult

    End Function

    Protected Function GenerateObject(ByRef AType As Type) As Object

        Dim LTemp As Type = Nothing

        If FTypeBind.TryGetValue(AType, LTemp) Then AType = LTemp
        Return Activator.CreateInstance(AType)

    End Function

    Protected Function ReadArray(ByVal AType As Type) As Object

        Dim LType As Type = GetType(Object)
        Dim LValues As IList
        Dim LIsList As Boolean

        If AType.IsArray OrElse Not GetType(IList).IsAssignableFrom(AType) Then
            If AType.IsArray Then LType = AType.GetElementType
            LValues = New List(Of Object)
            LIsList = False
        Else
            If AType.GetGenericArguments.Length > 0 Then LType = AType.GetGenericArguments(0)
            LValues = CType(GenerateObject(AType), IList)
            LIsList = True
        End If

        CheckNext("["c)
        FLevel += 1
        If ReadNext(ReadOption.Hold) <> "]"c Then
            Do
                LValues.Add(ReadValue(LType))
                If Not ReadWhenExpected(","c) Then Exit Do
            Loop
        End If
        FLevel -= 1
        CheckNext("]"c)

        If LIsList Then
            Return LValues
        Else
            Return GenerateArray(LType, LValues)
        End If

    End Function

    Protected Function ReadObject(ByVal AType As Type, ByVal AInstance As Object) As Object

        Dim LValue As Object
        Dim LValueType As Type = GetType(Object)
        Dim LDictionary As IDictionary = Nothing
        Dim LStructure As ValueType = Nothing
        Dim LMember As ObjectMember

        If AInstance Is Nothing Then AInstance = GenerateObject(AType)
        If AType.IsClass Then
            LDictionary = TryCast(AInstance, IDictionary)
            If LDictionary IsNot Nothing AndAlso AType.GetGenericArguments.Length > 1 Then LValueType = AType.GetGenericArguments()(1)
        Else
            LStructure = CType(AInstance, ValueType)
        End If

        CheckNext("{"c)
        FLevel += 1
        If ReadNext(ReadOption.Hold) <> "}"c Then
            Do
                If LDictionary IsNot Nothing Then
                    LDictionary.Add(ReadPropertyName, ReadValue(LValueType))
                Else
                    LMember = ObjectMember.Create(AType, ReadPropertyName)
                    If LMember Is Nothing Then
                        ReadValue(GetType(Object))
                    Else
                        LValue = ReadNullableValue(LMember.MemberType)
                        If LStructure IsNot Nothing Then
                            LMember.SetValue(LStructure, LValue)
                        Else
                            LMember.SetValue(AInstance, LValue)
                        End If
                    End If
                End If
                If Not ReadWhenExpected(","c) Then Exit Do
            Loop
        End If
        FLevel -= 1
        CheckNext("}"c)

        If Not AType.IsClass Then AInstance = LStructure
        Return AInstance

    End Function

    Protected Function ReadValue(ByVal AType As Type, Optional ByVal AInstance As Object = Nothing) As Object

        Try
            Select Case Char.ToLower(ReadNext(ReadOption.Hold))
                Case "n"c
                    If Not ReadNonString.SameText(NULL) Then Throw New FormatException("Die Zeichenfolge wurde nicht als gültiger Nullwert erkannt.")
                    Return Nothing

                Case "f"c, "t"c
                    Return Boolean.Parse(ReadNonString)

                Case "+"c, "-"c, "0"c To "9"c
                    Return Convert.ChangeType(Decimal.Parse(ReadNonString, FCulture), AType, FCulture)

                Case """"c
                    Return DecodeStringValue(AType, ReadString)

                Case "["c
                    Return ReadArray(AType)

                Case "{"c
                    Return ReadObject(AType, AInstance)

                Case Else
                    UnexpectedCharFound()
                    Return Nothing

            End Select
        Catch ex As DeserializationException
            Throw
        Catch ex As Exception
            Throw New DeserializationException(Me, ex, "Eine Instanz vom Typ '{0}' konnte nicht deserialisiert werden.", AType.ToString)
        End Try

    End Function

    Protected Function ReadNullableValue(ByVal AType As Type) As Object

        If AType.IsGenericType AndAlso AType.GetGenericTypeDefinition.Equals(GetType(Nullable(Of ))) Then
            Return ReadValue(Nullable.GetUnderlyingType(AType))
        Else
            Return ReadValue(AType)
        End If

    End Function

    Public Function InternalDeserialize(ByVal AInput As String, ByVal AType As Type, ByVal AInstance As Object) As Object

        FLevel = 0
        FInput = AInput
        FInputIndex = 0
        FComment = False
        Return ReadValue(AType, AInstance)

    End Function

    Public Function Deserialize(Of T)(ByVal AInput As String) As T

        Return CType(InternalDeserialize(AInput, GetType(T), Nothing), T)

    End Function

    Public Sub Deserialize(Of T)(ByVal AInput As String, ByRef AOutput As T)

        AOutput = Deserialize(Of T)(AInput)

    End Sub

    Public Sub DeserializeInstance(ByVal AInput As String, ByVal AInstance As Object)

        InternalDeserialize(AInput, AInstance.GetType, AInstance)

    End Sub

    Public Shared Function DirectDeserialize(Of T)(ByVal AInput As String) As T

        With New JSONSerializer()
            Return .Deserialize(Of T)(AInput)
        End With

    End Function

    Public Shared Sub DirectDeserialize(Of T)(ByVal AInput As String, ByRef AOutput As T)

        With New JSONSerializer()
            .Deserialize(AInput, AOutput)
        End With

    End Sub

    Public Shared Sub DirectDeserializeInstance(ByVal AInput As String, ByVal AInstance As Object)

        With New JSONSerializer()
            .DeserializeInstance(AInput, AInstance)
        End With

    End Sub

#End Region

#Region " Interne Klassen "

    Public Class DeserializationException
        Inherits Exception

        Private FPosition As Integer

        Public ReadOnly Property Position() As Integer
            Get
                Return FPosition
            End Get
        End Property

        Public Sub New(ByVal sender As JSONSerializer, ByVal message As String, ByVal ParamArray args As Object())

            Me.New(sender, Nothing, message, args)

        End Sub

        Public Sub New(ByVal sender As JSONSerializer, ByVal innerException As Exception, ByVal message As String, ByVal ParamArray args As Object())

            MyBase.New(String.Format(message, args), innerException)
            FPosition = sender.FInputIndex

        End Sub

    End Class

    Protected Class ObjectMember

        Private FField As FieldInfo
        Private FProperty As PropertyInfo

        Public ReadOnly Property MemberType() As Type
            Get
                If FField IsNot Nothing Then
                    Return FField.FieldType
                Else
                    Return FProperty.PropertyType
                End If
            End Get
        End Property

        Public Sub New(ByVal AField As FieldInfo, ByVal AProperty As PropertyInfo)

            FField = AField
            FProperty = AProperty

        End Sub

        Public Shared Function Create(ByVal AType As Type, ByVal AName As String) As ObjectMember

            Const BINDING_FLAGS As BindingFlags = BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase

            Dim LField As FieldInfo
            Dim LProperty As PropertyInfo

            LField = AType.GetField(AName, BINDING_FLAGS)
            If LField IsNot Nothing Then Return New ObjectMember(LField, Nothing)

            LProperty = AType.GetProperty(AName, BINDING_FLAGS)
            If LProperty IsNot Nothing Then Return New ObjectMember(Nothing, LProperty)

            Return Nothing

        End Function

        Public Sub SetValue(ByVal AInstance As Object, ByVal AValue As Object)

            If FField IsNot Nothing Then
                FField.SetValue(AInstance, AValue)
            Else
                FProperty.SetValue(AInstance, AValue, Nothing)
            End If

        End Sub

        Public Sub SetValue(ByVal AInstance As ValueType, ByVal AValue As Object)

            If FField IsNot Nothing Then
                FField.SetValue(AInstance, AValue)
            Else
                FProperty.SetValue(AInstance, AValue, Nothing)
            End If

        End Sub

    End Class

#End Region

End Class