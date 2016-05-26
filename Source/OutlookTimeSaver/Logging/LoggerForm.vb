Public Class LoggerForm

    Private m_LoggingEventsBindingSource As New BindingSource

    Public Sub New()

        Dim dgvCellStyleTimestamp = New DataGridViewCellStyle()
        dgvCellStyleTimestamp.Format = "HH:mm:ss:fff"

        InitializeComponent()

        With m_LoggingEventsBindingSource
            .DataSource = LoggerFormAppender.LoggingEvents
            .AllowNew = False
        End With

        With dgvLog
            .DataSource = m_LoggingEventsBindingSource
            .MultiSelect = False
            .AutoResizeColumns()
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False

            .Columns("LoggerName").Visible = False
            .Columns("LocationInformation").Visible = False
            .Columns("Repository").Visible = False
            .Columns("Timestamp").DefaultCellStyle = dgvCellStyleTimestamp
        End With

    End Sub

    Public Sub SetLayoutPosition()

        Me.SuspendLayout()
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        Me.Height = CInt(Screen.PrimaryScreen.WorkingArea.Height * 0.2)

        Me.Top = CInt(Screen.PrimaryScreen.WorkingArea.Height * 0.8)
        Me.Left = 0
        Me.ResumeLayout()

        Me.TopMost = True

    End Sub

    Private Delegate Sub RefreshDataCallback()

    Public Sub RefreshData()

        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If dgvLog.InvokeRequired Then
            Dim d As New RefreshDataCallback(AddressOf RefreshData)
            Me.Invoke(d)
            Return
        End If

        m_LoggingEventsBindingSource.ResetBindings(False)
        dgvLog.AutoResizeColumns()

    End Sub

    Private Sub ClearToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        dgvLog.Rows.Clear()
    End Sub
End Class