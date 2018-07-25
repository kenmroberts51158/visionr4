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
Public Class frmSchemaView

    Public filename As String = ""
    Public connectionString As String = ""
    Public schemaType As VisionHelper.schemaType

    Private Sub frmSchemaView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Me.filename & " Schema " & VisionHelper.getSchemaTypeText(Me.schemaType)
        VisionHelper.getSchemaInformation(Me.DataGridView1, Me.schemaType, Me.connectionString, Me.filename)
    End Sub

    Private Sub frmSchemaView_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.DataGridView1.Width = Me.Width - 10
        Me.DataGridView1.Height = Me.Height - 103
    End Sub

    Private Sub exitApplication()
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.exitApplication()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.exitApplication()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        VisionHelper.excelExport(Me.DataGridView1)
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        VisionHelper.exportToWord(Me.DataGridView1)
    End Sub

    Private Sub ExcelExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExcelExportToolStripMenuItem.Click
        VisionHelper.excelExport(Me.DataGridView1)
    End Sub

    Private Sub WordExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordExportToolStripMenuItem.Click
        VisionHelper.exportToWord(Me.DataGridView1)
    End Sub

End Class