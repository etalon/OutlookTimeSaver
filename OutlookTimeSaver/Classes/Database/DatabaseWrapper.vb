Imports System.Data.Common
Imports System.IO

''' <summary>
''' Dieser Wrapper soll den Datenbankzugriff auf SQLite, MySQL und MSSQL vereinheitlichen.
''' Die Konfiguration kann entweder aus einer Konfigurationsdatei gelesen werden (so kann das DBMS per Konfig gewechselt werden),
''' oder der ConnectionString kann überladen werden (vgl. SQlite-Translationdatenbank).
''' Ich habe mich dagegen entschieden diesen Wrapper auf dem MDE zur Verfügung zu stellen, da wir auf dem MDE nur Sqlite haben
''' und hier zudem Klassen und Aufrufe benötigt werden, welche nicht im IpClient existieren.
''' </summary>
''' <remarks></remarks>
Public MustInherit Class DatabaseWrapper
    Implements IDisposable

#Region "Klassen & Variablen"

    Public Enum eDbType
        Unknown
        MySql
        Sqlite
        MSSQL
    End Enum

    Public Class Parameter
        Public Shared DefaultDatabaseWrapperConfigOnNewConfigFile As _EtiDatabaseWrapperConfig
    End Class

    Public Class _EtiDatabaseWrapperConfig

        Public DbType As String = leDbType.Name(eDbType.Unknown)
        Public EnableSqlLogging As Boolean = True

        Public MySqlConfig As New _MySqlConfig
        Public SqliteConfig As New _SqliteConfig
        Public MSSQLConfig As New _MSSQLConfig

        Public Class _MySqlConfig
            Public Hostname As String = ""
            Public Username As String = ""
            Public Password As String = ""
            Public Database As String = ""
        End Class

        Public Class _SqliteConfig
            Public FileName As String = ""
            Public Version As Integer = 3
            Public UseUTF16Encoding As Boolean = True

            Public Function GetAbsoluteFileName(p_ConfigFileName As String) As String

                Return EtiPath.RelativeToAbsFile(FileName, IO.Path.GetDirectoryName(p_ConfigFileName))

            End Function

        End Class

        Public Class _MSSQLConfig
            Public Hostname As String = ""
            Public Username As String = ""
            Public Password As String = ""
            Public Database As String = ""
        End Class

    End Class

    ''' <summary>
    ''' ConnectionString, der manuell gesetzt wird und nicht aus der Konfiguration kommt
    ''' </summary>
    ''' <remarks></remarks>
    Protected m_ConnectionString As String = ""

    Protected m_Connection As DbConnection
    Protected m_Command As DbCommand
    Protected m_Transaction As DbTransaction
    Protected m_TransactionCount As Integer
    Protected m_RollbackOccured As Boolean
    Protected m_DataReader As IDataReader
    Protected m_DbType As eDbType

    Protected m_DatabaseWrapperConfig As _EtiDatabaseWrapperConfig
    Protected m_DatabaseWrapperConfigFile As String = ""

    Public Shared leDbType As New clsLoadedEnum(GetType(eDbType))

#End Region

#Region "Verwaltung der Konfigurations-Objekte"

    ''' <summary>
    ''' Der Schlüssel ist der Dateipfad +name. In diesem Dictionary werden die Konfigurationsobjekte zwischengespeichert,
    ''' sodass diese nicht jedes Mal neu eingelesen werden müssen
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared m_DatabaseWrapperConfigDict As New Dictionary(Of String, _EtiDatabaseWrapperConfig)
    Private Shared m_SyncLockEtiDatabaseWrapperConfigDict As New Object
    Protected Shared m_ConfigFile As String

    Private Shared Function getEtiDatabaseWrapperConfig(ByRef p_ConfigFile As String) As _EtiDatabaseWrapperConfig

        Dim myConfig As _EtiDatabaseWrapperConfig = Nothing

        If String.IsNullOrEmpty(IO.Path.GetDirectoryName(p_ConfigFile)) Then
            p_ConfigFile = Path.Combine(AppDataPath, p_ConfigFile)
        End If

        m_ConfigFile = p_ConfigFile

        SyncLock m_SyncLockEtiDatabaseWrapperConfigDict

            If Not m_DatabaseWrapperConfigDict.TryGetValue(p_ConfigFile, myConfig) Then

                If Parameter.DefaultDatabaseWrapperConfigOnNewConfigFile Is Nothing Then
                    myConfig = New _EtiDatabaseWrapperConfig
                Else
                    myConfig = Parameter.DefaultDatabaseWrapperConfigOnNewConfigFile
                End If

                If Not File.Exists(p_ConfigFile) Then
                    WriteConfig(p_ConfigFile, myConfig)
                End If

                loadConfig(p_ConfigFile, myConfig)

                m_DatabaseWrapperConfigDict.Add(p_ConfigFile, myConfig)

            End If

        End SyncLock

        Return myConfig

    End Function

    Public ReadOnly Property EmbeddedDatabaseFilename() As String
        Get
            If Me.m_DbType <> eDbType.Sqlite Then
                Throw New NotSupportedException("Die aktuell verwendete Datenbank ist keine eingebettete Datenbank!")
            Else
                Return m_DatabaseWrapperConfig.SqliteConfig.GetAbsoluteFileName(m_ConfigFile)
            End If
        End Get
    End Property

#End Region

#Region "Quick-Statements"

    Public Class QuickStatements

        Protected m_Database As DatabaseWrapper

        Public Sub New(ByVal database As DatabaseWrapper)

            m_Database = database

        End Sub

        Protected Function getParamString(ByVal count As Integer) As String

            Static Dim memory As New Dictionary(Of Integer, String)
            Dim result As String = Nothing
            Dim temp As System.Text.StringBuilder

            If Not memory.TryGetValue(count, result) Then
                temp = New System.Text.StringBuilder(count * 3)
                For i = 0 To count - 1
                    If temp.Length > 0 Then temp.Append(","c)
                    temp.Append("@" & i.ToString)
                Next
                result = temp.ToString
                memory.Add(count, result)
            End If
            Return result

        End Function

        ''' <summary>
        ''' Fügt einen neuen Datensatz hinzu. Es müssen für alle Felder der Tabelle die Daten in der richtigen Reihenfolge angegeben werden.
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="args"></param>
        ''' <remarks></remarks>
        Public Overridable Sub Insert(ByVal tableName As String, ByVal ParamArray args As Object())

            m_Database.Execute(String.Format("INSERT INTO {0} VALUES ({1})", tableName, getParamString(args.Count)), args)

        End Sub

    End Class

    Private m_QuickStatements As QuickStatements = Nothing

    Protected Overridable Function CreateQuickStatements() As QuickStatements

        Return New QuickStatements(Me)

    End Function

    ''' <summary>
    ''' Ermöglicht die einfache Ausführung häufig benötiger Aufgaben.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Quick() As QuickStatements
        Get
            If m_QuickStatements Is Nothing Then m_QuickStatements = CreateQuickStatements()
            Return m_QuickStatements
        End Get
    End Property

#End Region

#Region "Konstruktoren & Co."

    ''' <summary>
    ''' Standard-Konstruktor, der auf die Standard-Datenbank-Einstellung zurückgreift
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateInstance() As DatabaseWrapper
        Return CreateInstance("", "", eDbType.Unknown)
    End Function

    Public Shared Function CreateInstance(Optional ByVal p_ConfigFile As String = "") As DatabaseWrapper
        Return CreateInstance(p_ConfigFile, "", eDbType.Unknown)
    End Function

    Public Shared Function CreateInstance(ByVal p_OverloadedConnectionString As String, ByVal p_OverloadedConnectionType As eDbType) As DatabaseWrapper
        Return CreateInstance("", p_OverloadedConnectionString, p_OverloadedConnectionType)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="p_OverloadedConnectionString">Überladener Connection-String, falls man die Konfigurationsdatei nicht nutzen möchte</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CreateInstance(ByVal p_ConfigFile As String, ByVal p_OverloadedConnectionString As String, ByVal p_OverloadedConnectionType As eDbType) As DatabaseWrapper

        Dim myObject As DatabaseWrapper
        Dim myConfig As _EtiDatabaseWrapperConfig

        If String.IsNullOrEmpty(p_ConfigFile) Then
            p_ConfigFile = "EtiDatabaseWrapperConfig.cfg"
        End If

        myConfig = getEtiDatabaseWrapperConfig(p_ConfigFile)

        If p_OverloadedConnectionType = eDbType.Unknown Then
            p_OverloadedConnectionType = DirectCast(leDbType.Index(myConfig.DbType), eDbType)
        End If

        Select Case p_OverloadedConnectionType
            Case eDbType.MSSQL
                myObject = New DatabaseMSSQL
                myObject.m_DbType = eDbType.MSSQL
            Case eDbType.MySql
                myObject = New DatabaseMySql
                myObject.m_DbType = eDbType.MySql
            Case eDbType.Sqlite
                myObject = New DatabaseSqlite()
                myObject.m_DbType = eDbType.Sqlite
            Case Else
                Throw New Exception(String.Format("Es wurde kein gültiger Datenbanktyp in der Datei '{0}' festgelegt!", p_ConfigFile))
        End Select

        With myObject
            .m_ConnectionString = p_OverloadedConnectionString
            .m_DatabaseWrapperConfig = myConfig
            .m_DatabaseWrapperConfigFile = p_ConfigFile
            .Init()
        End With

        Return myObject

    End Function

    Friend Shared Sub WriteConfig(ByVal p_FileName As String, ByVal p_Object As _EtiDatabaseWrapperConfig)
        File.WriteAllText(p_FileName, JSONSerializer.DirectSerialize(p_Object, True))
    End Sub

    Private Shared Sub loadConfig(ByVal p_FileName As String, ByRef p_Object As _EtiDatabaseWrapperConfig)
        p_Object = JSONSerializer.DirectDeserialize(Of _EtiDatabaseWrapperConfig)(File.ReadAllText(p_FileName))
    End Sub

    Protected Sub writeConfig()
        WriteConfig(m_DatabaseWrapperConfigFile, m_DatabaseWrapperConfig)
    End Sub

    Protected Sub loadConfig()
        loadConfig(m_DatabaseWrapperConfigFile, m_DatabaseWrapperConfig)
    End Sub

    Protected Friend MustOverride Sub Init()

#End Region

#Region "Transaktionen"

    Public Sub StartTransaction(Optional ByVal iso As IsolationLevel = IsolationLevel.RepeatableRead)

        addToLog(String.Format("Transaction #{0}: Start ({1})", m_TransactionCount + 1, iso.ToString))

        ' TODO: Wieso wird bei SQLite kein Isolation Level übergeben?
        ' TODO: Wenn die Nested Transaction ein anderes Isolation Level will, muss ein Fehler ausgelöst werden.

        If m_TransactionCount = 0 Then
            Select Case m_DbType
                Case eDbType.Sqlite
                    m_Transaction = m_Connection.BeginTransaction()

                Case Else
                    m_Transaction = m_Connection.BeginTransaction(iso)

            End Select
            m_Command.Transaction = m_Transaction
            m_RollbackOccured = False
        End If
        m_TransactionCount += 1

    End Sub

    Protected Sub StopTransaction()

        addToLog(String.Format("Transaction #{0}: {1}", m_TransactionCount, If(m_RollbackOccured, "Rollback", "Commit")))

        m_TransactionCount -= 1
        If m_TransactionCount = 0 Then
            If m_RollbackOccured Then
                m_Transaction.Rollback()
            Else
                m_Transaction.Commit()
            End If
            m_Command.Transaction = Nothing
            m_Transaction = Nothing
        End If

    End Sub

    Public Sub CommitTransaction()

        StopTransaction()

    End Sub

    Public Sub RollbackTransaction()

        m_RollbackOccured = True
        StopTransaction()

    End Sub

#End Region

#Region "Excecutes"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="SqlStatement">Parameter mit @0, @1, @2 etc...</param>
    ''' <param name="aArgs" >Parameter für die Platzhalter (@0, @1, @2 etc...)</param>
    ''' <returns>Anzahl aktualisierter Datensätze</returns>
    ''' <remarks></remarks>
    Public Function ExecuteNonQuery(ByVal SqlStatement As String, ByVal ParamArray aArgs As Object()) As Integer

        initializeCommand(SqlStatement, aArgs)

        If m_DatabaseWrapperConfig.EnableSqlLogging Then
            'addToLog("DB.ExecuteNonQuery(): " & m_Command.CommandTextWithReplacedParameters, clsLog.DataFlow.ToDatabase)
        End If

        Return m_Command.ExecuteNonQuery()

    End Function

    Public Function Execute(ByVal sqlStatement As String, ByVal ParamArray args As Object()) As Integer

        Return ExecuteNonQuery(sqlStatement, args)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="SqlStatement">Parameter mit @0, @1, @2 etc...</param>
    ''' <param name="aArgs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(Of T)(ByVal SqlStatement As String, ByVal ParamArray aArgs As Object()) As T

        initializeCommand(SqlStatement, aArgs)

        If m_DatabaseWrapperConfig.EnableSqlLogging Then
            'addToLog("DB.ExecuteScalar(): " & m_Command.CommandTextWithReplacedParameters, clsLog.DataFlow.ToDatabase)
        End If

        ' Anmerkung: Sollte CType sein und nicht DirectCast.
        ' Hintergrund: COUNT(*) gibt unter MySql einen Long-Wert zurück und
        ' dieser kann mit DirectCast nicht in einen Integer umgewandelt werden
        Return CType(m_Command.ExecuteScalar(), T)

    End Function

    ''' <summary>
    ''' Liest einen Wert vom Typ T aus der Datenbank. DBNull wird dabei zu Nothing umgewandelt.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="sqlStatement"></param>
    ''' <param name="args"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadScalar(Of T)(ByVal sqlStatement As String, ByVal ParamArray args As Object()) As T
        Return ReadScalarDefault(Of T)(sqlStatement, Nothing, args)
    End Function

    Public Function ReadScalarDefault(Of T)(ByVal sqlStatement As String, ByVal defaultValue As T, ByVal ParamArray args As Object()) As T

        Dim result As Object

        result = ExecuteScalar(Of Object)(sqlStatement, args)
        If TypeOf result Is DBNull Then result = defaultValue
        Return CType(result, T)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="SqlStatement">Parameter mit @0, @1, @2 etc...</param>
    ''' <param name="aArgs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteReader(ByVal SqlStatement As String, ByVal ParamArray aArgs As Object()) As IDataReader

        If m_DataReader IsNot Nothing AndAlso Not m_DataReader.IsClosed Then
            Throw New Exception("Vorheriger DataReader wurde nicht geschlossen!")
        End If

        initializeCommand(SqlStatement, aArgs)

        If m_DatabaseWrapperConfig.EnableSqlLogging Then
            'm_Log.Add("DB.ExecuteReader(): " & m_Command.CommandTextWithReplacedParameters, clsLog.DataFlow.ToDatabase)
        End If

        m_DataReader = m_Command.ExecuteReader()

        Return m_DataReader

    End Function

    Public Function ReadRows(ByVal sqlStatement As String, ByVal ParamArray args As Object()) As List(Of FetchedDataRow)

        Dim result As New List(Of FetchedDataRow)

        Using reader = ExecuteReader(sqlStatement, args)
            While reader.Read
                result.Add(New FetchedDataRow(reader))
            End While
        End Using
        Return result

    End Function

    Public Function TryReadSingleRow(ByRef row As FetchedDataRow, ByVal sqlStatement As String, ByVal ParamArray args As Object()) As Boolean

        Dim rows As List(Of FetchedDataRow)

        rows = ReadRows(sqlStatement, args)
        Select Case rows.Count
            Case 0
                Return False

            Case 1
                row = rows.First
                Return True

            Case Else
                Throw New Exception("Es wurden mehrere Datenbankeinträge gefunden, aber nur einer erwartet.")

        End Select

    End Function

    Public Function ReadSingleRow(ByVal sqlStatement As String, ByVal ParamArray args As Object()) As FetchedDataRow

        Dim row As FetchedDataRow = Nothing

        If Not TryReadSingleRow(row, sqlStatement, args) Then
            Throw New Exception("Der Datenbankeintrag konnte nicht geladen werden.")
        End If

        Return row

    End Function

    Public Overridable Function LastInsertId() As Integer

        Throw New NotSupportedException()

    End Function

#End Region

#Region "DataTable & DataSet"

    Public MustOverride Function GetDataTable(ByVal sql As String, ByRef da As IDataAdapter, ByRef dt As DataTable, ByRef bs As BindingSource, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs As Object()) As DataTable

    Public Function GetDataTable(ByVal sql As String, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs As Object()) As DataTable

        Return GetDataTable(sql, Nothing, Nothing, Nothing, False, aArgs)

    End Function

    Public Function GetDataTable(ByVal sql As String, ByVal ParamArray aArgs As Object()) As DataTable

        Return GetDataTable(sql, False, aArgs)

    End Function

    Public Function ReadTable(ByVal sqlStatement As String, ByVal ParamArray args As Object()) As DataTable

        Return GetDataTable(sqlStatement, args)

    End Function

    Public Function ReadTableByName(ByVal tableName As String) As DataTable

        Return ReadTable(String.Format("SELECT * FROM {0}", tableName))

    End Function

    Public MustOverride Function GetDataSet(ByVal sql As String, ByRef da As IDataAdapter, ByRef ds As DataSet, ByRef bs As BindingSource, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs As Object()) As System.Data.DataSet

    ''' <summary>
    ''' Führt ein DataAdapter.Update aus. Diese Funktion wird benötigt, da man beim IDataAdapter nur ein DataSet übergeben kann, nicht aber eine DataTable
    ''' </summary>
    ''' <param name="p_Da"></param>
    ''' <param name="p_Dt"></param>
    ''' <remarks></remarks>
    Public MustOverride Function DataAdapterUpdate(ByVal p_Da As IDataAdapter, ByVal p_Dt As DataTable) As Integer

#End Region

#Region "Schema"

    Public Function GetTables() As List(Of String)

        Dim columns As New List(Of String)

        For Each oRow As DataRow In m_Connection.GetSchema("Tables").Rows
            columns.Add(oRow("TABLE_NAME").ToString.ToLower)
        Next

        Return columns

    End Function

    Public Function GetTableColumns(ByVal tableName As String) As List(Of String)

        Dim columns As New List(Of String)

        For Each oRow As DataRow In m_Connection.GetSchema("Columns").Rows
            ' falls es sich um die gesuchte Tabelle handelt, 
            ' jetzt den Feldnamen ausgeben
            If oRow("TABLE_NAME").ToString = tableName Then
                columns.Add(oRow("COLUMN_NAME").ToString.ToLower)
            End If
        Next

        Return columns

    End Function

#End Region

#Region "Hilfsfunktionen"

    Protected Sub addToLog(ByVal text As String)

        If m_DatabaseWrapperConfig.EnableSqlLogging Then
            Log.Info(text)
        End If

    End Sub

    <Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities")>
    Protected Sub initializeCommand(ByVal sSQL As String, ByVal ParamArray aArgs As Object())

        m_Command.CommandText = sSQL
        m_Command.Parameters.Clear()

        For i As Integer = 0 To aArgs.Length - 1
            m_Command.Parameters.Add(getNewDbParameter(i.ToString, aArgs(i)))
        Next

    End Sub

    Protected MustOverride Function getNewDbParameter(ByVal parameterName As String, ByVal value As Object) As DbParameter

#End Region

#Region "Dispose-Support"

    Private disposedValue As Boolean = False        ' So ermitteln Sie überflüssige Aufrufe

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                ' Versuchen einen offenen DataReader zu schließen
                If m_DataReader IsNot Nothing Then
                    If Not m_DataReader.IsClosed Then
                        m_DataReader.Close()
                        m_DataReader.Dispose()
                    End If
                End If

                ' Wenn wir noch eine Transaktion haben
                If m_Transaction IsNot Nothing Then
                    If m_Connection.State = ConnectionState.Open Then
                        RollbackTransaction()
                    End If
                End If

                ' Verbindung schließen
                If m_Connection.State = ConnectionState.Open Then
                    m_Connection.Close()
                    m_Connection.Dispose()
                    m_Connection = Nothing
                End If

            End If


        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#End Region

    ''' <summary>
    ''' Diese Prozedur überprüft ob die Datenbankverbindung offen ist und öffnet diese ggf. wieder.
    ''' Wenn die Verbindung zu lange offen ist (24h) kann es passieren
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub VerifyConnectionState()

        If m_Connection.State <> ConnectionState.Open Then
            m_Connection.Open()
        End If

    End Sub

End Class
