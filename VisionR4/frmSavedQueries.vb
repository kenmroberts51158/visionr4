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
Public Class frmSavedQueries

    Private myDataTable As DataTable

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ListView1.Font = My.Settings.MyFont
        Me.setListViewMenuItems()
        Me.refreshMe()
    End Sub

    Private Sub setListViewMenuItems()
        Me.ListView1.View = My.Settings.ListView

        Me.LargeIconsToolStripMenuItem.Checked = False
        Me.SmallIconsToolStripMenuItem.Checked = False
        Me.ListToolStripMenuItem.Checked = False
        Me.TileToolStripMenuItem.Checked = False

        If Me.ListView1.View = View.LargeIcon Then Me.LargeIconsToolStripMenuItem.Checked = True
        If Me.ListView1.View = View.SmallIcon Then Me.SmallIconsToolStripMenuItem.Checked = True
        If Me.ListView1.View = View.List Then Me.ListToolStripMenuItem.Checked = True
        If Me.ListView1.View = View.Tile Then Me.TileToolStripMenuItem.Checked = True
    End Sub

    Private Sub refreshMe()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.fillMyDataTable()
        Me.fillList()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub Form4_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.myDataTable.Dispose()
            Me.myDataTable = Nothing
        Catch ex As Exception
            ' just smile and wave boys, smile and wave.
        End Try
    End Sub

    Private Sub fillMyDataTable()
        Dim v As New vision()
        myDataTable = v.SavedQueries
        Dim s As New visionTableAdapters.SavedQueriesTableAdapter
        s.Fill(myDataTable)
        v.Dispose()
        v = Nothing
    End Sub

    Private Sub fillList()
        Me.ListView1.Items.Clear()
        For i As Integer = 0 To Me.myDataTable.Rows.Count - 1
            Dim t As New ListViewItem
            t.ImageIndex = 0
            t.Text = Me.myDataTable.Rows(i).Item("queryName")
            Me.ListView1.Items.Add(t)
        Next
    End Sub

    Private Sub deleteThis()
        If IsNothing(ListView1.FocusedItem) = False Then
            If MsgBox("Are you sure you want to delete " & ListView1.FocusedItem.Text & "?", MsgBoxStyle.OkCancel, Application.ProductName) = MsgBoxResult.Ok Then
                Dim dr As DataRow = Me.myDataTable.Rows.Find(ListView1.FocusedItem.Text)
                If IsNothing(dr) = False Then
                    Dim s As New visionTableAdapters.SavedQueriesTableAdapter
                    s.Delete(ListView1.FocusedItem.Text)
                    s.Dispose()
                    s = Nothing
                    Me.refreshMe()
                End If
            End If
        Else
            MsgBox("Nothing selected to delete.", MsgBoxStyle.Information, My.Application.Info.Title)
        End If
    End Sub

    Private Sub runUserSQL(ByVal sender As Object)
        Dim temp As String = ""
        For i As Integer = 0 To Me.myDataTable.Rows.Count - 1
            If sender.FocusedItem.Text = Me.myDataTable.Rows(i).Item("queryName") Then
                temp = Me.myDataTable.Rows(i).Item("querySQL")
                Dim s As String = ""
                Dim p As Integer = -1
                If InStr(Me.myDataTable.Rows(i).Item("querySQL"), "[parameter]") Then
                    Dim parms() As String = Me.myDataTable.Rows(i).Item("queryParms").ToString.Split(";")
                    For x As Integer = 0 To parms.Length - 1
                        p = InStr(parms(x), "[parameter]")
                        s = parms(x).Replace("[parameter]", InputBox(Mid(parms(x), 1, p - 2), "Parameter Prompt", ""))
                        temp = temp.Replace(parms(x), s)
                    Next
                End If
                Dim frmUpdateSQL As New frmUserSQL
                frmUpdateSQL.MdiParent = Me.MdiParent
                frmUpdateSQL.mySQL = temp
                frmUpdateSQL.dsn = Me.myDataTable.Rows(i).Item("queryDSN")
                frmUpdateSQL.Show()
                Exit For
            End If
        Next
    End Sub

    Private Sub Form4_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.ListView1.Width = Me.Width - 10
        Me.ListView1.Height = Me.Height - 95
    End Sub

    Private Sub exitApplication()
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.exitApplication()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.exitApplication()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.deleteThis()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.refreshMe()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Me.refreshMe()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Me.deleteThis()
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        Me.runUserSQL(sender)
    End Sub

    Private Sub LargeIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LargeIconsToolStripMenuItem.Click
        Me.ListView1.View = View.LargeIcon
        My.Settings.ListView = Me.ListView1.View
        Me.LargeIconsToolStripMenuItem.Checked = True
        Me.SmallIconsToolStripMenuItem.Checked = False
        Me.ListToolStripMenuItem.Checked = False
        Me.TileToolStripMenuItem.Checked = False
    End Sub

    Private Sub SmallIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SmallIconsToolStripMenuItem.Click
        Me.ListView1.View = View.SmallIcon
        My.Settings.ListView = Me.ListView1.View
        Me.LargeIconsToolStripMenuItem.Checked = False
        Me.SmallIconsToolStripMenuItem.Checked = True
        Me.ListToolStripMenuItem.Checked = False
        Me.TileToolStripMenuItem.Checked = False
    End Sub

    Private Sub ListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListToolStripMenuItem.Click
        Me.ListView1.View = View.List
        My.Settings.ListView = Me.ListView1.View
        Me.LargeIconsToolStripMenuItem.Checked = False
        Me.SmallIconsToolStripMenuItem.Checked = False
        Me.ListToolStripMenuItem.Checked = True
        Me.TileToolStripMenuItem.Checked = False
    End Sub

    Private Sub TileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TileToolStripMenuItem.Click
        Me.ListView1.View = View.Tile
        My.Settings.ListView = Me.ListView1.View
        Me.LargeIconsToolStripMenuItem.Checked = False
        Me.SmallIconsToolStripMenuItem.Checked = False
        Me.ListToolStripMenuItem.Checked = False
        Me.TileToolStripMenuItem.Checked = True
    End Sub
End Class