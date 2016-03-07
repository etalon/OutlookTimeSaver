Imports System.Data.SQLite

Public Class DatabaseSqlite
    Inherits DatabaseWrapper

    Public Shared Shadows Function CreateInstance(ByVal fileName As String) As DatabaseWrapper
        Return DatabaseWrapper.CreateInstance(GetConnectionString(fileName), eDbType.Sqlite)
    End Function

    Public Shared Shadows Function CreateInstance(ByVal p_FileName As String, ByVal p_UseExistingDatabaseFile As Boolean) As DatabaseWrapper

        Dim fi As IO.FileInfo

        If p_UseExistingDatabaseFile Then
            fi = New IO.FileInfo(p_FileName)

            If Not fi.Exists Then
                Throw New Exception("Datenbankdatei '" & p_FileName & "' existiert nicht!")
            End If

            If fi.Length = 0 Then
                Throw New Exception("Datenbankdatei '" & p_FileName & "' ist ungültig!")
            End If
        Else
            IO.File.Delete(p_FileName)
        End If

        Return CreateInstance(p_FileName)

    End Function


    Protected Friend Overrides Sub Init()

        If String.IsNullOrEmpty(m_ConnectionString) Then
            With m_DatabaseWrapperConfig.SqliteConfig
                m_ConnectionString = GetConnectionString(.GetAbsoluteFileName(m_ConfigFile), .Version, .UseUTF16Encoding)
            End With
        End If

        m_Connection = New SQLiteConnection(m_ConnectionString)
        m_Connection.Open()

        m_Command = m_Connection.CreateCommand()

        m_Command.CommandText = "PRAGMA foreign_keys = ON;"
        m_Command.ExecuteNonQuery()

    End Sub

    Public Shared Function GetConnectionString(ByVal fileName As String, Optional ByVal version As Integer = 3, Optional ByVal useUtf16Encoding As Boolean = True) As String

        With New SQLiteConnectionStringBuilder
            .DataSource = EtiPath.RelativeToAbsFile(fileName)
            .Version = version
            .UseUTF16Encoding = useUtf16Encoding
            Return .ToString
        End With

    End Function

    Public Overrides Function LastInsertId() As Integer

        Return ReadScalar(Of Integer)("SELECT last_insert_rowid()")

    End Function

    Public Overrides Function GetDataTable(ByVal sql As String, ByRef da As System.Data.IDataAdapter, ByRef dt As System.Data.DataTable, ByRef bs As System.Windows.Forms.BindingSource, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs() As Object) As System.Data.DataTable

        Dim dataAdapter As SQLiteDataAdapter = Nothing

        initializeCommand(sql, aArgs)

        dataAdapter = New SQLiteDataAdapter(DirectCast(m_Command, SQLite.SQLiteCommand))

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        If activateCommandBuilder Then
            With New SQLiteCommandBuilder(dataAdapter)
                .ConflictOption = System.Data.ConflictOption.OverwriteChanges
                .SetAllValues = True

                If sql.Contains(" JOIN ") Then
                    dataAdapter.DeleteCommand = .GetDeleteCommand()
                    dataAdapter.UpdateCommand = .GetUpdateCommand()
                    dataAdapter.InsertCommand = .GetInsertCommand()
                End If
            End With
        End If

        dt = New DataTable
        dataAdapter.Fill(dt)

        bs = New BindingSource
        bs.DataSource = dt

        da = dataAdapter
        Return dt

    End Function

    Public Overrides Function GetDataSet(ByVal sql As String, ByRef da As System.Data.IDataAdapter, ByRef ds As System.Data.DataSet, ByRef bs As System.Windows.Forms.BindingSource, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs() As Object) As System.Data.DataSet
        Throw New NotSupportedException
    End Function

    Public Overrides Function DataAdapterUpdate(ByVal p_Da As System.Data.IDataAdapter, ByVal p_Dt As System.Data.DataTable) As Integer
        Return DirectCast(p_Da, SQLiteDataAdapter).Update(p_Dt)
    End Function

    Protected Overrides Function getNewDbParameter(ByVal parameterName As String, ByVal value As Object) As System.Data.Common.DbParameter
        Return New SQLiteParameter(parameterName, value)
    End Function

#Region "Helper"

    ''' <summary>
    ''' Gibt das angegebene Datum als String für das Speichern in einer SQLite-Stringspalte zurück. SQLite hat keine Datumsspalte.
    ''' </summary>
    ''' <param name="myDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSqliteDate(ByVal myDate As Date) As String
        ' Muss so eingebaut sein, weil bei der Datumsformatierung immer wieder ein Punkt zwischen den Zeiten erscheint.
        Return Replace(myDate.ToString("yyyy-MM-dd HH:mm:ss"), ".", ":")
    End Function

    ''' <summary>
    ''' Ersetzt einige Nicht-SQLite konforme Befehle mit SQLite-Syntax
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <remarks></remarks>
    Public Shared Sub TranslateCommandsToSqlite(ByRef sql As String)

        sql = sql.Replace("auto_increment", "AUTOINCREMENT")
        sql = sql.Replace("IDENTITY (1,1)", "AUTOINCREMENT")
        sql = sql.Replace("IDENTITY(1,1)", "AUTOINCREMENT")
        sql = sql.Replace("datetime", "nvarchar(27)")

    End Sub

#End Region

End Class
