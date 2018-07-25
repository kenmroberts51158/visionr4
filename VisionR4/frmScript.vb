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
Public Class frmScript

    Public selectedDSN As String = ""
    Private visionUser As New VisionUser()

    Private Sub frmScript_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtUserSQL.SelectionStart = 0
        VisionHelper.getDataSourceNames(Me.cboDSNs)
        If selectedDSN.Length > 0 Then
            Me.cboDSNs.SelectedIndex = Me.cboDSNs.FindString(selectedDSN)
        End If
    End Sub

    Private Sub frmScript_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.txtUserSQL.Width = Me.Width - 35
        Me.txtUserSQL.Height = Me.Height - 157
        Me.ProgressBar1.Width = Me.Width - 35
    End Sub

    Private Sub runScript()
        If Me.txtUserSQL.Text.Trim.Length = 0 Then
            MsgBox("I would be happy to run your SQL script if you had one for me to run.", MsgBoxStyle.Information, My.Application.Info.Title)
            Exit Sub
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = 10
        Try
            Dim scriptSqls() As String = Me.txtUserSQL.Text.Split(";")
            Me.ProgressBar1.Maximum = scriptSqls.Length - 1
            For i As Integer = 0 To scriptSqls.Length - 1
                If scriptSqls(i).Trim.Length > 0 Then
                    Me.ProgressBar1.Increment(1)
                    VisionHelper.executeNonQuery(scriptSqls(i), Me.getConnectionString)
                End If
            Next
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox(ex.Message, MsgBoxStyle.Critical, Application.ProductName)
        End Try
        Me.ProgressBar1.Value = 0
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub copyToClipBoard()
        If Me.txtUserSQL.Text.Trim.Length > 0 Then
            Clipboard.Clear()
            Clipboard.SetText(Me.txtUserSQL.Text, TextDataFormat.Text)
        End If
    End Sub

    Private Sub saveToFile()
        Me.SaveFileDialog1.Title = "Save Script"
        Me.SaveFileDialog1.Filter = "SQL files (*.sql)|*.sql|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(SaveFileDialog1.FileName)
            sw.Write(Me.txtUserSQL.Text)
            sw.Close()
        End If
    End Sub

    Private Sub openFromFile()
        Me.OpenFileDialog1.FileName = ""
        Me.OpenFileDialog1.Title = "Open Script"
        Me.OpenFileDialog1.Filter = "SQL files (*.sql)|*.sql|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        If Me.OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim sr As New System.IO.StreamReader(OpenFileDialog1.FileName)
            Me.txtUserSQL.Text = sr.ReadToEnd
            sr.Close()
        End If
    End Sub

    Private Function getConnectionString() As String
        Dim rtnValue As String = ""
        rtnValue = "DSN=" & Me.cboDSNs.SelectedItem & ";"
        rtnValue &= "UID=" & Me.txtUserId.Text & ";"
        rtnValue &= "PWD=" & Me.txtPassword.Text & ";"
        Return rtnValue
    End Function

    Private Sub OpenFromFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenFromFileToolStripMenuItem.Click
        Me.openFromFile()
    End Sub

    Private Sub cboDSNs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDSNs.SelectedIndexChanged
        Me.visionUser.getDefaults(cboDSNs.SelectedItem.ToString)
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CopyToClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToClipboardToolStripMenuItem.Click
        Me.copyToClipBoard()
    End Sub

    Private Sub SaveToFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToFileToolStripMenuItem.Click
        Me.saveToFile()
    End Sub

    Private Sub RunScriptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunScriptToolStripMenuItem.Click
        Me.runScript()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Me.runScript()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.runScript()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.copyToClipBoard()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.saveToFile()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.openFromFile()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.Close()
    End Sub
End Class