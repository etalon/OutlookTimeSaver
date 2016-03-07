Imports System.Data.SqlClient

Public Class DatabaseMSSQL
    Inherits DatabaseWrapper

    Protected Friend Overrides Sub Init()

        Dim decryptedPassword As String = ""
        Dim encryptionDone As Boolean

        If String.IsNullOrEmpty(m_ConnectionString) Then
            With m_DatabaseWrapperConfig.MSSQLConfig
                decryptedPassword = BinEncodeDecode.GetDecryptedPassword(.Password, encryptionDone)
                m_ConnectionString = GetConnectionString(.Hostname, .Username, decryptedPassword, .Database)
            End With
        End If

        If encryptionDone Then
            writeConfig()
        End If

        m_Connection = New SqlConnection(m_ConnectionString)
        m_Connection.Open()

        m_Command = m_Connection.CreateCommand

    End Sub

    Public Shared Function GetConnectionString(ByVal hostname As String, ByVal username As String, ByVal password As String, ByVal database As String) As String

        With New SqlConnectionStringBuilder
            .DataSource = hostname
            .UserID = username
            .Password = password
            .InitialCatalog = database
            Return .ToString
        End With

    End Function

    Public Overrides Function GetDataTable(ByVal sql As String, ByRef da As System.Data.IDataAdapter, ByRef dt As System.Data.DataTable, ByRef bs As System.Windows.Forms.BindingSource, ByVal activateCommandBuilder As Boolean, ByVal ParamArray aArgs() As Object) As System.Data.DataTable

        Dim dataAdapter As SqlDataAdapter

        initializeCommand(sql, aArgs)
        dataAdapter = New SqlDataAdapter(DirectCast(m_Command, SqlCommand))

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        If activateCommandBuilder Then
            With New SqlCommandBuilder(dataAdapter)
                .ConflictOption = System.Data.ConflictOption.OverwriteChanges
                .SetAllValues = True

                If sql.Contains("JOIN") Then
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

        Dim dataAdapter As SqlDataAdapter

        initializeCommand(sql, aArgs)
        dataAdapter = New SqlDataAdapter(DirectCast(m_Command, SqlCommand))

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        If activateCommandBuilder Then
            With New SqlCommandBuilder(dataAdapter)
                .ConflictOption = System.Data.ConflictOption.OverwriteChanges
                .SetAllValues = True

                If sql.Contains("JOIN") Then
                    dataAdapter.DeleteCommand = .GetDeleteCommand()
                    dataAdapter.UpdateCommand = .GetUpdateCommand()
                    dataAdapter.InsertCommand = .GetInsertCommand()
                End If
            End With
        End If

        ds = New DataSet
        da.Fill(ds)

        bs = New BindingSource
        bs.DataSource = ds

        da = dataAdapter
        Return ds

    End Function

    Public Overrides Function DataAdapterUpdate(ByVal p_Da As System.Data.IDataAdapter, ByVal p_Dt As System.Data.DataTable) As Integer
        Return DirectCast(p_Da, SqlDataAdapter).Update(p_Dt)
    End Function

    Protected Overrides Function getNewDbParameter(ByVal parameterName As String, ByVal value As Object) As System.Data.Common.DbParameter
        Return New SqlParameter(parameterName, value)
    End Function

End Class
