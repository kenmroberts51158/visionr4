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
Public Class frmUserSQL
    Public mySQL As String = ""
    Public filename As String = ""
    Public dsn As String = ""
    Private visionUser As New VisionUser()

    Public WriteOnly Property User() As VisionUser
        Set(ByVal value As VisionUser)
            Me.visionUser = value
        End Set
    End Property

    Private Sub frmUpdateSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VisionHelper.getDataSourceNames(Me.cboDSNs)
        If Me.dsn.Trim.Length > 0 Then
            For i As Integer = 0 To Me.cboDSNs.Items.Count - 1
                If Me.cboDSNs.Items.Item(i) = Me.dsn Then
                    Me.cboDSNs.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        Me.txtUserSQL.Text = mySQL
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
        Me.txtUserSQL.Font = My.Settings.MyFont

    End Sub

    Protected Overrides Sub Finalize()
        Me.visionUser = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub frmUserSQL_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.txtUserSQL.Width = Me.Width - 35
        Me.txtUserSQL.Height = Me.Height - 175
    End Sub

    Private Function getConnectionString() As String
        Dim rtnValue As String = ""
        rtnValue = "DSN=" & Me.cboDSNs.SelectedItem & ";"
        rtnValue &= "UID=" & Me.txtUserId.Text & ";"
        rtnValue &= "PWD=" & Me.txtPassword.Text & ";"
        Return rtnValue
    End Function

    Private Sub runSQL()
        Try
            If Me.txtUserSQL.Text.Trim.Length > 0 Then
                If (Me.txtUserSQL.Text.TrimStart.StartsWith("select", StringComparison.CurrentCultureIgnoreCase) = True) _
                    Or (chkAllowNonSelect.Checked = True) Then
                    Dim f As New frmResults
                    f.MdiParent = Me.MdiParent
                    f.filename = Me.filename
                    f.connectionString = Me.getConnectionString()
                    f.dsn = Me.cboDSNs.SelectedItem
                    f.sql = Me.txtUserSQL.Text
                    Me.visionUser.UserID = Me.txtUserId.Text
                    Me.visionUser.Password = Me.txtPassword.Text
                    f.User = Me.visionUser
                    f.Show()
                Else
                    VisionHelper.executeNonQuery(Me.txtUserSQL.Text, Me.getConnectionString)
                End If
            Else
                MsgBox("I would be happy to run your SQL statement if you had one for me to run.", MsgBoxStyle.Information, My.Application.Info.Title)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        End Try
    End Sub

    Private Sub saveQuery()
        Dim queryName As String = InputBox("Please enter a Query Name: ", "Save Query", "")
        If queryName <> "" Then
            Try
                Dim s As New visionTableAdapters.SavedQueriesTableAdapter
                s.Insert(queryName, Me.cboDSNs.SelectedItem, Me.txtUserSQL.Text, "")
                s.Dispose()
                s = Nothing
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
            End Try
        End If
    End Sub

    Private Sub copyToClipBoard()
        Clipboard.Clear()
        Clipboard.SetText(Me.txtUserSQL.Text, TextDataFormat.Text)
    End Sub

    Private Sub outputTemplateText(ByVal whatTemplate As String)
        Dim sb As New System.Text.StringBuilder
        sb.AppendLine("")
        Select Case whatTemplate.ToUpper
            Case "UPDATE"
                sb.AppendLine("UPDATE tablename")
                sb.AppendLine("SET colunm = 'value'")
                sb.AppendLine("WHERE colunm = 'value'")
            Case "INSERT"
                sb.AppendLine("INSERT INTO tablename")
                sb.AppendLine("[(columnname, ...)]")
                sb.AppendLine("VALUES (constant, ...)")
            Case "JOIN"
                sb.AppendLine("/* JOIN or INNER JOIN - An operation that retrieves rows from multiple ")
                sb.AppendLine("source tables by comparing the values from columns shared between ")
                sb.AppendLine("the source tables. An inner join excludes rows from a source table ")
                sb.AppendLine("that have no matching rows in the other source tables. */ ")
                sb.AppendLine("")
                sb.AppendLine("SELECT * FROM tablename1")
                sb.AppendLine("JOIN tablename2 ON (tablename1.field1 = tablename2.field1)")
                sb.AppendLine("WHERE tablename1.field1 = 'value' /* etc, etc */")
            Case "OUTER JOIN"
                sb.AppendLine("/* OUTER JOIN - A join that includes all the rows from the joined ")
                sb.AppendLine("tables that meet the search conditions, even rows from one table ")
                sb.AppendLine("for which there is no matching row in the other join table. For ")
                sb.AppendLine("result set rows returned when a row in one table is not matched ")
                sb.AppendLine("by a row from the other table, a value of NULL is supplied for ")
                sb.AppendLine("all result set columns that are resolved to the table that had ")
                sb.AppendLine("the missing row. */")
                sb.AppendLine("")
                sb.AppendLine("SELECT * FROM tablename1")
                sb.AppendLine("OUTER JOIN tablename2 ON (tablename1.field1 = tablename2.field1)")
                sb.AppendLine("WHERE tablename1.field1 = 'value' /* etc, etc */")
            Case "LEFT OUTER JOIN"
                sb.AppendLine("/* LEFT OUTER JOIN - The result of a left outer join for tables ")
                sb.AppendLine("A and B always contains all records of the 'left' table (A), even ")
                sb.AppendLine("if the join-condition does not find any matching record in the ")
                sb.AppendLine("'right' table (B). This means that if the ON clause matches 0 (zero) ")
                sb.AppendLine("records in B, the join will still return a row in the result — but ")
                sb.AppendLine("with NULL in each column from B. */")
                sb.AppendLine("")
                sb.AppendLine("SELECT * FROM tablenameA")
                sb.AppendLine("LEFT OUTER JOIN tablenameB ON (tablenameA.field1 = tablenameB.field1)")
                sb.AppendLine("WHERE tablenameA.field1 = 'value' /* etc, etc */")
            Case "RIGHT OUTER JOIN"
                sb.AppendLine("/* RIGHT OUTER JOIN - A right outer join closely resembles a left outer ")
                sb.AppendLine("join, except with the tables reversed. Every record from the 'right' ")
                sb.AppendLine("table (B) will appear in the joined table at least once. If no matching")
                sb.AppendLine("row from the 'left' table (A) exists, NULL will appear in columns ")
                sb.AppendLine("from A for those records that have no match in A. */")
                sb.AppendLine("")
                sb.AppendLine("SELECT * FROM tablenameA")
                sb.AppendLine("RIGHT OUTER JOIN tablenameB ON (tablenameA.field1 = tablenameB.field1)")
                sb.AppendLine("WHERE tablenameA.field1 = 'value' /* etc, etc */")
        End Select
        sb.AppendLine("")
        Me.txtUserSQL.Text += sb.ToString()
    End Sub

    Private Sub exitApplication()
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.exitApplication()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.copyToClipBoard()
    End Sub

    Private Sub RunToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunToolStripMenuItem.Click
        Me.runSQL()
    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Me.runSQL()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.exitApplication()
    End Sub

    Private Sub CopyToClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToClipboardToolStripMenuItem.Click
        Me.copyToClipBoard()
    End Sub

    Private Sub SaveQueryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveQueryToolStripMenuItem.Click
        Me.saveQuery()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.saveQuery()
    End Sub

    Private Sub txtUserId_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserId.GotFocus
        Me.txtUserId.SelectAll()
    End Sub

    Private Sub txtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassword.GotFocus
        Me.txtPassword.SelectAll()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.runSQL()
    End Sub

    Private Sub UpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub InsertToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub JoinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JoinToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub OuterJoinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OuterJoinToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub LeftOuterJoinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeftOuterJoinToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub RightOuterJoinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RightOuterJoinToolStripMenuItem.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub UpdateToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub InsertToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub JoinToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JoinToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub OuterJoinToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OuterJoinToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub LeftOuterJoinToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeftOuterJoinToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub RightOuterJoinToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RightOuterJoinToolStripMenuItem1.Click
        outputTemplateText(sender.text)
    End Sub

    Private Sub cboDSNs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDSNs.SelectedIndexChanged
        Me.visionUser.getDefaults(cboDSNs.SelectedItem.ToString)
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub

End Class