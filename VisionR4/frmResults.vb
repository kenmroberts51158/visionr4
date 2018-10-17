'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*
'-*                                                    *
'-* Application  . . : Vision R4                       *
'-*                                                    *
'-* Copyright  . . . : 2010 Kamikaze Software.         *
'-*                    Carver Ma, 02330  USA           *
'-*                    All rights reserved.            *
'-*                                                    *
'-* Provided as-is with no warranties,                 * 
'-* your mileage may vary.                             *
'-*                                                    *
'-* Portions developed by Piso Mojado Software.        *
'-*                                                    *
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*
'-*
Imports myODbc = System.Data.Odbc
Imports System.ComponentModel

Public Class frmResults
    Public filename As String = ""
    Public connectionString As String = ""
    Public sql As String = ""
    Private rowIndex As Integer = -1
    Private WithEvents myDataGridDataTable As DataTable
    Private myCurrentDataRowData As DataGridViewRow
    Public dsn As String = ""
    Private visionUser As New VisionUser()

    Public WriteOnly Property User() As VisionUser
        Set(ByVal value As VisionUser)
            Me.visionUser = value
        End Set
    End Property

    Private Sub frmResults2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Results: " & Me.dsn & " " & Me.filename
        Me.refreshData()
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Me.myDataGridDataTable.Dispose()
            Me.myDataGridDataTable = Nothing
            Me.myCurrentDataRowData = Nothing
            MyBase.Finalize()
        Catch ex As Exception
            ' just smile and wave boys, smile and wave!
        End Try
    End Sub

    Private Sub ExitApplication()
        Me.Close()
    End Sub

    Private Sub refreshData()
        Try
            Me.DataGridView1.CancelEdit()
        Catch ex As Exception
            'Cancel Edit cannot be called while in the onRowChange event. 
            'This occurs when an error happens trying to insert a record that can't be inserted for some reason (nulls, etc.)
            'Don't fail, just let the data refresh happen.

        End Try
        Me.showHideControls(False, True)
        Me.BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub showHideControls(ByVal showControls As Boolean, ByVal showButton As Boolean)
        If IsNothing(Me.filename) OrElse Me.filename.Trim.Length = 0 Then
            Me.ToolStripButton2.Enabled = False
            Me.ToolStripButton6.Enabled = False
            Me.ToolStripButton8.Enabled = False
            Me.ViewSchemaToolStripMenuItem.Enabled = False
            Me.ViewKeysToolStripMenuItem.Enabled = False
            Me.XMLExportToolStripMenuItem.Enabled = False
            Me.DataGridView1.AllowUserToAddRows = False
            Me.DataGridView1.AllowUserToDeleteRows = False
            Me.DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
        End If

        Me.DataGridView1.Visible = showControls
        Me.ToolStrip1.Visible = showControls
        Me.MenuStrip1.Visible = showControls
        Me.ContextMenuStrip1.Enabled = showControls
        Me.Button1.Enabled = showButton
        Me.Button1.Visible = showButton
        Me.Label1.Visible = showButton
    End Sub

    Private Sub frmResults2_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.DataGridView1.Width = Me.Width - 25
        Me.DataGridView1.Height = Me.Height - 100
    End Sub

    Private Sub viewSchema(ByVal schemaType As VisionHelper.schemaType)
        Dim f As New frmSchemaView
        f.MdiParent = Me.MdiParent
        f.filename = Me.filename
        f.connectionString = Me.connectionString
        f.schemaType = schemaType
        f.Show()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim bw As BackgroundWorker = CType(sender, BackgroundWorker)
        e.Result = VisionHelper.getDataTable(Me.filename, Me.sql, Me.connectionString)
        If bw.CancellationPending Then
            e.Cancel = True
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled = False Then
            Me.DataGridView1.DataSource = e.Result
            Me.myDataGridDataTable = e.Result
        End If
        Me.showHideControls(True, False)
    End Sub

    Private Sub handleContextMenuClick(ByVal action As String)
        Dim tempSql As String = ""
        Select Case action
            Case "update"
                tempSql = VisionHelper.getUpdateSQL(Me.DataGridView1, Me.rowIndex, Me.myCurrentDataRowData)
            Case "delete"
                tempSql = VisionHelper.getDeleteSQL(Me.DataGridView1, Me.rowIndex)
            Case "insert"
                tempSql = VisionHelper.getInsertSQL(Me.DataGridView1, Me.rowIndex)
            Case Else
                tempSql = Me.sql
        End Select

        Dim frmUpdateSQL As New frmUserSQL
        frmUpdateSQL.MdiParent = Me.MdiParent
        frmUpdateSQL.mySQL = tempSql
        frmUpdateSQL.filename = Me.filename
        frmUpdateSQL.dsn = Me.dsn
        frmUpdateSQL.User = Me.visionUser
        frmUpdateSQL.Show()
    End Sub

    Private Sub xmlExport()
        Try
            Dim dt As DataTable = CType(Me.DataGridView1.DataSource, DataTable)
            Dim path As String = InputBox("Enter file path: (ie. C:\temp\myXml.xml)", "XML Save", "")
            If path = "" Then Exit Sub
            dt.WriteXml(path)
            dt = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.BackgroundWorker1.CancelAsync()
        Me.BackgroundWorker2.CancelAsync()
        Me.BackgroundWorker3.CancelAsync()
        Me.ExitApplication()
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim bw As BackgroundWorker = CType(sender, BackgroundWorker)
        VisionHelper.excelExport(Me.DataGridView1)
        If bw.CancellationPending Then
            e.Cancel = True
        End If
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Me.DataGridView1.Visible = True
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim bw As BackgroundWorker = CType(sender, BackgroundWorker)
        VisionHelper.exportToWord(Me.DataGridView1)
        If bw.CancellationPending Then
            e.Cancel = True
        End If
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Me.DataGridView1.Visible = True
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.DataGridView1.Visible = False
        Me.BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.viewSchema(VisionHelper.schemaType.Columns)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.refreshData()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.ExitApplication()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.DataGridView1.Visible = False
        Me.BackgroundWorker3.RunWorkerAsync()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.viewSchema(VisionHelper.schemaType.Indexes)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.ExitApplication()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MsgBox(e.Exception.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        Me.DataGridView1.CancelEdit()
    End Sub

    Private Sub DataGridView1_RowEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
        Me.rowIndex = e.RowIndex

        Dim x As System.Data.DataRowView = CType(Me.DataGridView1.Rows(Me.rowIndex).DataBoundItem, System.Data.DataRowView)

        Try
            Me.myCurrentDataRowData = Me.DataGridView1.Rows(Me.rowIndex).Clone
            For i As Integer = 0 To Me.DataGridView1.Rows(Me.rowIndex).Cells.Count - 1
                Me.myCurrentDataRowData.Cells(i).Value = Me.DataGridView1.Rows(Me.rowIndex).Cells.Item(i).Value
                Me.myCurrentDataRowData.Cells(i).ValueType = Me.DataGridView1.Rows(Me.rowIndex).Cells.Item(i).ValueType
            Next
        Catch ex As Exception
            Me.myCurrentDataRowData = Nothing
        End Try
    End Sub

    Private Sub myDataGridDataTable_RowChanging(sender As Object, e As System.Data.DataRowChangeEventArgs) Handles myDataGridDataTable.RowChanging
        Try
            Select Case e.Action
                Case DataRowAction.Add
                    VisionHelper.executeNonQuery(VisionHelper.getInsertSQL(Me.DataGridView1, Me.rowIndex), Me.connectionString)
                Case DataRowAction.Change
                    VisionHelper.executeNonQuery(VisionHelper.getUpdateSQL(Me.DataGridView1, Me.rowIndex, Me.myCurrentDataRowData), Me.connectionString)
            End Select
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & "When you click OK your results will be refreshed.", MsgBoxStyle.Critical, My.Application.Info.Title)
            Me.refreshData()
        End Try

    End Sub

    Private Sub myDataGridDataTable_RowDeleting(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles myDataGridDataTable.RowDeleting
        Try
            VisionHelper.executeNonQuery(VisionHelper.getDeleteSQL(Me.DataGridView1, Me.rowIndex), Me.connectionString)
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & "When you click OK your results will be refreshed.", MsgBoxStyle.Critical, My.Application.Info.Title)
            Me.refreshData()
        End Try
    End Sub

    Private Sub InsertSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertSQLToolStripMenuItem.Click
        Me.handleContextMenuClick("insert")
    End Sub

    Private Sub UpdateSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateSQLToolStripMenuItem.Click
        Me.handleContextMenuClick("update")
    End Sub

    Private Sub DeleteSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteSQLToolStripMenuItem.Click
        Me.handleContextMenuClick("delete")
    End Sub

    Private Sub CurrentSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentSQLToolStripMenuItem.Click
        Me.handleContextMenuClick("")
    End Sub

    Private Sub ExcelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcelToolStripMenuItem.Click
        Me.DataGridView1.Visible = False
        Me.BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub WordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordToolStripMenuItem.Click
        Me.DataGridView1.Visible = False
        Me.BackgroundWorker3.RunWorkerAsync()
    End Sub

    Private Sub ViewSchemaToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewSchemaToolStripMenuItem.Click
        Me.viewSchema(VisionHelper.schemaType.Columns)
    End Sub

    Private Sub ViewKeysToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewKeysToolStripMenuItem.Click
        Me.viewSchema(VisionHelper.schemaType.Indexes)
    End Sub

    Private Sub RefreshToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Me.refreshData()
    End Sub

    Private Sub CurrentSQLToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentSQLToolStripMenuItem1.Click
        Me.handleContextMenuClick("")
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Me.handleContextMenuClick("")
    End Sub

    Private Sub XMLExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XMLExportToolStripMenuItem.Click
        Me.xmlExport()
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Me.xmlExport()
    End Sub

End Class