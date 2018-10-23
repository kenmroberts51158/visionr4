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
Public Class frmNewQuery
    Private fieldsTable As DataTable
    Private indexesTable As DataTable
    Private visionUser As New VisionUser()

    Private Sub frmNewQuery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VisionHelper.getDataSourceNames(Me.cboDSNs)
        Me.cboBool.SelectedIndex = 0
        Me.cboAndOr.SelectedIndex = 0
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub

    Private Sub frmNewQuery_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.fieldsTable = Nothing
            Me.indexesTable = Nothing
            Me.visionUser = Nothing
        Catch ex As Exception
            ' ;-p
        End Try
    End Sub

    Private Function getConnectionString() As String
        Dim rtnValue As String = ""
        rtnValue = "DSN=" & Me.cboDSNs.SelectedItem & ";"
        rtnValue &= "UID=" & Me.txtUserId.Text & ";"
        rtnValue &= "PWD=" & Me.txtPassword.Text & ";"
        Return rtnValue
    End Function

    Private Function buildSQL(ByVal showPrompt As Boolean, ByVal isCount As Boolean) As String
        Dim sql As String = ""

        If Me.ckbDistinct.Checked = False Then
            sql = "SELECT "
        Else
            sql = "SELECT DISTINCT "
        End If

        If isCount = False Then
            If Me.ListBoxSelectedFields.Items.Count = 0 Then
                sql &= "* "
            Else
                For i As Integer = 0 To Me.ListBoxSelectedFields.Items.Count - 1
                    If Me.ListBoxSelectedFields.Items(i).Trim.Contains(" ") = False Then
                        sql &= Me.ListBoxSelectedFields.Items(i) & ","
                    Else
                        sql &= "[" & Me.ListBoxSelectedFields.Items(i) & "],"
                    End If
                Next
                sql = sql.TrimEnd(",")
            End If
        Else
            sql &= "count(*) "
        End If

        ''If Me.cboTables.SelectedItem.ToString.Trim.Contains(" ") = False Then
        If CType(Me.cboTables.SelectedItem, TableListItem).TableName.Trim.Contains(" ") = False Then
            sql &= " FROM " & CType(Me.cboTables.SelectedItem, TableListItem).TableName
        Else
            sql &= " FROM [" & CType(Me.cboTables.SelectedItem, TableListItem).TableName & "]"
        End If

        If Me.ListBoxWhereClause.Items.Count > 0 Then
            sql &= " WHERE "
            For i As Integer = 0 To Me.ListBoxWhereClause.Items.Count - 1
                Dim p As Integer = InStr(Me.ListBoxWhereClause.Items(i), "[parameter]")
                If showPrompt = False Then p = 0
                If p > 0 Then
                    Dim temp As String = InputBox(Mid(Me.ListBoxWhereClause.Items(i), 1, p - 2), "Parameter Prompt")
                    sql &= Me.ListBoxWhereClause.Items(i).ToString.Replace("[parameter]", temp)
                Else
                    sql &= Me.ListBoxWhereClause.Items(i)
                End If
            Next
        End If

        If sql.EndsWith("AND ") Or sql.EndsWith("OR ") Then
            sql = sql.Remove(sql.Length - 4, 4)
        End If

        ' Add order by.
        If IsNothing(Me.cboOrderBy.SelectedItem) = False Then
            sql &= " ORDER BY " & Me.cboOrderBy.SelectedItem
            If Me.rdoAscending.Checked = True Then
                sql &= " ASC"
            Else
                sql &= " DESC"
            End If
        End If

        Return sql
    End Function

    Private Sub saveQuery()
        If IsNothing(Me.cboTables.SelectedItem) = False Then
            Dim queryName As String = InputBox("Please enter a Query Name: ", "Save Query", "")
            If queryName <> "" Then
                Try
                    Dim s As New visionTableAdapters.SavedQueriesTableAdapter
                    s.Insert(queryName, Me.cboDSNs.SelectedItem, Me.buildSQL(False, False), Me.getParameterStrings)
                    s.Dispose()
                    s = Nothing
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
                End Try
            End If
        Else
            MsgBox("No Table selected. Try clicking the Connect button.", MsgBoxStyle.Information, My.Application.Info.Title)
        End If
    End Sub

    Private Function getParameterStrings()
        Dim rtnValue As String = ""
        For i As Integer = 0 To Me.ListBoxWhereClause.Items.Count - 1
            If InStr(Me.ListBoxWhereClause.Items(i), "[parameter]") Then
                rtnValue &= Me.ListBoxWhereClause.Items(i)
                rtnValue = rtnValue.Remove(rtnValue.Length - 4, 4)
                rtnValue &= ";"
            End If
        Next
        Return rtnValue.TrimEnd(";")
    End Function

    Private Sub ExitApplication()
        Me.Close()
    End Sub

    Private Sub clearCbosAndLists()
        Me.cboTables.Items.Clear()
        Me.cboTables.Refresh()
        Me.ListBoxFields.Items.Clear()
        Me.ListBoxSelectedFields.Items.Clear()
        Me.cboSelectedField.Items.Clear()
        Me.cboSearchValue.Text = ""
        Me.ListBoxWhereClause.Items.Clear()
        Me.cboOrderBy.Items.Clear()
        Me.cboOrderBy.Text = ""
        Me.btnRunQuery.Enabled = False
    End Sub

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        clearCbosAndLists()
        Me.visionUser.UserID = Me.txtUserId.Text
        Me.visionUser.Password = Me.txtPassword.Text
        Dim dt As DataTable = VisionHelper.getSchemaDataTable(VisionHelper.schemaType.Tables, Me.getConnectionString)
        If IsNothing(dt) = False Then
            For i As Integer = 0 To dt.Rows.Count - 1
                ''Me.cboTables.Items.Add(dt.Rows(i).Item("TABLE_NAME").ToString)
                Dim itm As New TableListItem(dt.Rows(i).Item("TABLE_NAME").ToString, dt.Rows(i).Item("REMARKS").ToString)
                Me.cboTables.Items.Add(itm)
                itm = Nothing
            Next
            Me.cboTables.SelectedIndex = 0
            dt.Dispose()
            dt = Nothing
            Me.btnRunQuery.Enabled = True
        End If
    End Sub

    Private Sub cboTables_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTables.SelectedValueChanged
        Me.ListBoxFields.Items.Clear()
        Me.ListBoxSelectedFields.Items.Clear()
        Me.ListBoxWhereClause.Items.Clear()
        Me.cboSelectedField.Items.Clear()
        Me.cboOrderBy.Items.Clear()
        Me.cboOrderBy.Text = ""
        Me.cboSearchValue.Text = ""
        Me.fieldsTable = VisionHelper.getSchemaDataTable(VisionHelper.schemaType.Columns, Me.getConnectionString, CType(Me.cboTables.SelectedItem, TableListItem).TableName)
        Me.indexesTable = VisionHelper.getSchemaDataTable(VisionHelper.schemaType.Indexes, Me.getConnectionString, CType(Me.cboTables.SelectedItem, TableListItem).TableName)
        If IsNothing(fieldsTable) = False Then
            If fieldsTable.Rows.Count > 0 Then
                Dim tableSchema As String = fieldsTable.Rows(0).Item("TABLE_SCHEM").ToString()
                For i As Integer = 0 To fieldsTable.Rows.Count - 1
                    If tableSchema <> fieldsTable.Rows(i).Item("TABLE_SCHEM").ToString() Then Exit For
                    Me.ListBoxFields.Items.Add(fieldsTable.Rows(i).Item("COLUMN_NAME").ToString)
                    Me.cboSelectedField.Items.Add(fieldsTable.Rows(i).Item("COLUMN_NAME").ToString)
                    Me.cboOrderBy.Items.Add(fieldsTable.Rows(i).Item("COLUMN_NAME").ToString)
                Next
            End If
        End If
    End Sub

    Private Sub viewSchema(ByVal schemaType As VisionHelper.schemaType)
        If IsNothing(Me.cboTables.SelectedItem) = False Then
            Dim f As New frmSchemaView
            f.MdiParent = Me.MdiParent
            f.filename = CType(Me.cboTables.SelectedItem, TableListItem).TableName
            f.connectionString = Me.getConnectionString
            f.schemaType = schemaType
            f.Show()
        Else
            MsgBox("No Table selected. Try clicking the Connect button.", MsgBoxStyle.Information, My.Application.Info.Title)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.ListBoxSelectedFields.Items.AddRange(Me.ListBoxFields.Items)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.ListBoxFields.SelectedIndex > -1 Then
            Me.ListBoxSelectedFields.Items.Add(Me.ListBoxFields.SelectedItem)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Me.ListBoxSelectedFields.SelectedIndex > -1 Then
            Me.ListBoxSelectedFields.Items.RemoveAt(Me.ListBoxSelectedFields.SelectedIndex)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.ListBoxSelectedFields.Items.Clear()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ' moves items up the list
        If Me.ListBoxSelectedFields.SelectedIndex > 0 Then
            Dim oldSpot As Integer = Me.ListBoxSelectedFields.SelectedIndex
            Me.ListBoxSelectedFields.Items.Insert(oldSpot - 1, Me.ListBoxSelectedFields.SelectedItem)
            Me.ListBoxSelectedFields.Items.RemoveAt(oldSpot + 1)
            Me.ListBoxSelectedFields.SelectedIndex = (oldSpot - 1)
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ' moves items down the list
        If Me.ListBoxSelectedFields.SelectedIndex < Me.ListBoxSelectedFields.Items.Count - 1 Then
            Dim oldSpot As Integer = Me.ListBoxSelectedFields.SelectedIndex
            Me.ListBoxSelectedFields.Items.Insert(oldSpot + 2, Me.ListBoxSelectedFields.SelectedItem)
            Me.ListBoxSelectedFields.Items.RemoveAt(oldSpot)
            Me.ListBoxSelectedFields.SelectedIndex = (oldSpot + 1)
        End If
    End Sub

    Private Sub ListBoxFields_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxFields.Click
        If Me.ListBoxFields.SelectedIndex >= 0 Then
            Me.cboSelectedField.SelectedIndex = Me.ListBoxFields.SelectedIndex
        End If
    End Sub

    Private Sub ListBoxFields_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxFields.DoubleClick
        Try
            Me.ListBoxSelectedFields.Items.Add(Me.ListBoxFields.SelectedItem)
        Catch ex As Exception
            ' :-p
        End Try
    End Sub

    Private Sub ListBoxSelectedFields_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxSelectedFields.DoubleClick
        If Me.ListBoxSelectedFields.SelectedIndex >= 0 Then
            Me.ListBoxSelectedFields.Items.RemoveAt(Me.ListBoxSelectedFields.SelectedIndex)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.ExitApplication()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.ExitApplication()
    End Sub

    Private Sub btnAddSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSearch.Click
        If IsNothing(fieldsTable) = True Then Exit Sub
        If Len(Me.cboSelectedField.SelectedItem) = 0 Then Exit Sub

        Dim temp As String = ""
        If Me.cboSelectedField.SelectedItem.ToString.Trim.Contains(" ") = False Then
            temp = Me.cboSelectedField.SelectedItem & " " & Me.cboBool.Text & " "
        Else
            temp = "[" & Me.cboSelectedField.SelectedItem & "] " & Me.cboBool.Text & " "
        End If

        For i As Integer = 0 To fieldsTable.Rows.Count - 1
            If fieldsTable.Rows(i).Item("COLUMN_NAME").ToString = Me.cboSelectedField.SelectedItem Then
                If fieldsTable.Rows(i).Item("TYPE_NAME").ToString.ToUpper.Contains("CHAR") Then
                    temp &= "'" & Me.cboSearchValue.Text & "'"
                Else
                    temp &= Me.cboSearchValue.Text
                End If
                Exit For
            End If
        Next
        temp &= Me.cboAndOr.Text
        Me.ListBoxWhereClause.Items.Add(temp)

        'Put the cursor on the RunQuery button to enable enter to execute.
        btnRunQuery.Focus()
    End Sub

    Private Sub ListBoxWhereClause_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBoxWhereClause.Click
        If Me.ListBoxWhereClause.SelectedIndex >= 0 Then
            Me.ListBoxWhereClause.Items.RemoveAt(Me.ListBoxWhereClause.SelectedIndex)
        End If
    End Sub

    Private Sub btnRunQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRunQuery.Click
        Dim f As New frmResults
        f.MdiParent = Me.MdiParent
        f.filename = CType(Me.cboTables.SelectedItem, TableListItem).TableName
        f.connectionString = Me.getConnectionString()
        f.dsn = Me.cboDSNs.SelectedItem
        f.sql = Me.buildSQL(True, False)
        f.User = Me.visionUser
        f.Show()
    End Sub

    Private Sub btnUseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUseKey.Click
        If IsNothing(indexesTable) = True Then Exit Sub
        If indexesTable.Rows.Count = 0 Then
            MsgBox("Sorry, no Keys.", MsgBoxStyle.Information, My.Application.Info.Title)
            Exit Sub
        End If
        If IsNothing(fieldsTable) = True Then Exit Sub
        If IsNothing(Me.cboTables.SelectedItem) = True Then Exit Sub
        Me.ListBoxWhereClause.Items.Clear()

        Dim tableSchema As String = indexesTable.Rows(0).Item("TABLE_SCHEM").ToString()

        For i As Integer = 0 To indexesTable.Rows.Count - 1
            If tableSchema <> indexesTable.Rows(i).Item("TABLE_SCHEM").ToString() Then Exit For
            If indexesTable.Rows(i).Item("NON_UNIQUE") = 0 Then
                For y As Integer = 0 To fieldsTable.Rows.Count - 1
                    If fieldsTable.Rows(y).Item("COLUMN_NAME").ToString = indexesTable.Rows(i).Item("COLUMN_NAME").ToString Then
                        Dim temp As String = InputBox("Enter a value for " & fieldsTable.Rows(y).Item("COLUMN_NAME").ToString, "Use Key", "")
                        If temp.Trim.Length > 0 Then
                            If fieldsTable.Rows(y).Item("TYPE_NAME").ToString.ToUpper.Contains("CHAR") Then
                                temp = fieldsTable.Rows(y).Item("COLUMN_NAME").ToString & "= '" & temp & "' AND "
                            Else
                                temp = fieldsTable.Rows(y).Item("COLUMN_NAME").ToString & "= " & temp & " AND "
                            End If
                            Me.ListBoxWhereClause.Items.Add(temp)
                        End If
                        Exit For
                    End If
                Next
            End If
        Next

        'Put the cursor on the RunQuery button to enable enter to execute.
        btnRunQuery.Focus()
    End Sub

    Private Sub viewCurrentSQL()
        If IsNothing(Me.cboTables.SelectedItem) = False Then
            Dim frmUpdateSQL As New frmUserSQL
            frmUpdateSQL.MdiParent = Me.MdiParent
            frmUpdateSQL.mySQL = Me.buildSQL(False, False)
            frmUpdateSQL.filename = CType(Me.cboTables.SelectedItem, TableListItem).TableName
            frmUpdateSQL.User = Me.visionUser
            frmUpdateSQL.dsn = Me.cboDSNs.SelectedItem
            frmUpdateSQL.Show()
        Else
            MsgBox("No Table selected. Try clicking the Connect button.", MsgBoxStyle.Information, My.Application.Info.Title)
        End If
    End Sub

    Private Sub runRecordCount()
        If IsNothing(Me.cboTables.SelectedItem) = False Then
            Dim f As New frmResults
            f.MdiParent = Me.MdiParent
            f.filename = CType(Me.cboTables.SelectedItem, TableListItem).TableName
            f.connectionString = Me.getConnectionString()
            f.dsn = Me.cboDSNs.SelectedItem
            f.sql = Me.buildSQL(False, True)
            f.User = Me.visionUser
            f.Show()
        Else
            MsgBox("No Table selected. Try clicking the Connect button.", MsgBoxStyle.Information, My.Application.Info.Title)
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.viewSchema(VisionHelper.schemaType.Columns)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.viewSchema(VisionHelper.schemaType.Indexes)
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Me.viewSchema(VisionHelper.schemaType.Columns)
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Me.viewSchema(VisionHelper.schemaType.Indexes)
    End Sub

    Private Sub CurrentSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentSQLToolStripMenuItem.Click
        Me.viewCurrentSQL()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.viewCurrentSQL()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Me.saveQuery()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.saveQuery()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Me.runRecordCount()
    End Sub

    Private Sub RecordCountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecordCountToolStripMenuItem.Click
        Me.runRecordCount()
    End Sub

    Private Sub txtUserId_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserId.GotFocus
        Me.txtUserId.SelectAll()
    End Sub

    Private Sub txtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassword.GotFocus
        Me.txtPassword.SelectAll()
    End Sub

    Private Sub cboDSNs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDSNs.SelectedIndexChanged
        Me.visionUser.getDefaults(cboDSNs.SelectedItem.ToString)
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub

End Class