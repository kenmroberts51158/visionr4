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

Public Class VisionHelper
    Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Auto Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Integer, ByVal fDirection As Short, ByVal szDSN As String, ByVal cbDSNMax As Short, ByRef pcbDSN As Short, ByVal szDescription As String, ByVal cbDescriptionMax As Short, ByRef pcbDescription As Short) As Short
    Private Declare Auto Function SQLAllocEnv Lib "ODBC32.DLL" (ByRef env As Integer) As Short
    Private Const CMDTIMEDEFAULT As Integer = 60
    Private Shared cmdTimeOut As Integer = CMDTIMEDEFAULT

    Public Shared Property CommandTimeout() As Integer
        Get
            Return VisionHelper.cmdTimeOut
        End Get
        Set(ByVal value As Integer)
            VisionHelper.cmdTimeOut = value
        End Set
    End Property

    Public Shared Sub restoreCommandTimeoutToDefault()
        VisionHelper.CommandTimeout = VisionHelper.CMDTIMEDEFAULT
    End Sub

    Public Shared Function getDataTable(ByVal fileName As String, ByVal sql As String, ByVal connectionString As String) As DataTable
        Dim dt As DataTable = New DataTable(fileName)
        Try
            Dim da As myODbc.OdbcDataAdapter = New myODbc.OdbcDataAdapter(sql, connectionString)
            da.SelectCommand.CommandTimeout = VisionHelper.CommandTimeout
            da.Fill(dt)
            da.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        End Try
        Return dt
    End Function

    Public Shared Sub executeNonQuery(ByVal sql As String, ByVal connectionString As String)
        Dim conn As myODbc.OdbcConnection = New myODbc.OdbcConnection(connectionString)
        conn.Open()
        Dim cmd As myODbc.OdbcCommand = New myODbc.OdbcCommand(sql, conn)
        cmd.ExecuteNonQuery()

        cmd.Dispose()
        conn.Close()
        cmd = Nothing
        conn = Nothing
    End Sub

    Public Enum schemaType As Integer
        Columns = 1
        Tables = 2
        Indexes = 3
    End Enum

    Public Shared Function getSchemaTypeText(ByVal schemaType As VisionHelper.schemaType) As String
        Dim st As String = ""
        Select Case schemaType
            Case VisionHelper.schemaType.Columns
                st = "Columns"
            Case VisionHelper.schemaType.Indexes
                st = "Indexes"
            Case VisionHelper.schemaType.Tables
                st = "Tables"
        End Select
        Return st
    End Function

    Public Shared Sub getSchemaInformation(ByRef dgv As DataGridView, ByVal schemaType As VisionHelper.schemaType, ByVal connectionString As String, Optional ByVal fileName As String = Nothing)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Try
            Dim con As myODbc.OdbcConnection = New myODbc.OdbcConnection(connectionString)
            con.Open()
            Dim res(2) As String
            res(2) = fileName
            Dim table As DataTable = con.GetSchema(VisionHelper.getSchemaTypeText(schemaType), res)
            If schemaType = VisionHelper.schemaType.Columns Then
                Try
                    table = GetKnownColumnsTable(table)
                Catch ex As Exception
                    'No harm, no foul.
                End Try
            End If
            dgv.DataSource = table
            table.Dispose()
            con.Close()
            con = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Shared Function GetKnownColumnsTable(ByVal table As Data.DataTable) As Data.DataTable
        Dim newTable As New Data.DataTable(table.TableName)

        'Add columns
        Dim csList() As ColumnSchema = ColumnSchemaManager.GetColumnList()
        Dim cs As ColumnSchema
        For Each cs In csList
            newTable.Columns.Add(cs.ColumnName)
        Next

        'Copy the data
        Dim origRow As Data.DataRow
        For Each origRow In table.Rows()
            Dim nr As Data.DataRow = newTable.NewRow()
            For Each cs In csList
                nr(cs.ColumnName) = origRow(cs.ColumnName)
            Next
            newTable.Rows.Add(nr)
        Next

        Return newTable
    End Function

    Public Shared Function getSchemaDataTable(ByVal schemaType As VisionHelper.schemaType, ByVal connectionString As String, Optional ByVal fileName As String = Nothing) As DataTable
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim table As DataTable = Nothing
        Try
            Dim con As myODbc.OdbcConnection = New myODbc.OdbcConnection(connectionString)
            con.Open()
            Dim res(2) As String
            res(2) = fileName
            table = con.GetSchema(VisionHelper.getSchemaTypeText(schemaType), res)
            con.Close()
            con = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
        Return table
    End Function

    Public Shared Sub excelExport(ByRef dgv As DataGridView)
        Dim rowCountStop As Integer = 1
        If dgv.AllowUserToAddRows = True Then
            rowCountStop = 2
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Try
            Dim y As Integer = 1    ' excel row counter
            Dim xlApp As Object = CreateObject("excel.application")
            xlApp.Workbooks.Add()

            ' remove all but one worksheet
            Dim sc As Integer = xlApp.WorkSheets.Count
            xlApp.Worksheets.add()
            For i As Integer = sc To 1 Step -1
                xlApp.Worksheets(i).Delete()
            Next

            ' place field names in first row and bold them
            For i As Integer = 0 To dgv.ColumnCount - 1
                xlApp.Cells(y, i + 1).Value = dgv.Columns(i).HeaderText
                xlApp.Cells(y, i + 1).Font.Bold = True
            Next
            y += 1
            ' rest of data goes here.
            For i As Integer = 0 To dgv.Rows.Count - rowCountStop
                For x As Integer = 0 To dgv.ColumnCount - 1
                    If (dgv.Columns(x).ValueType.Name = "String") Then
                        xlApp.Cells(y, x + 1).NumberFormat = "@"
                    End If
                    xlApp.Cells(y, x + 1).Value = dgv.Rows(i).Cells(x).Value.ToString()
                Next
                y += 1

                If y > 65535 Then
                    ' this is Excel's limit. if hit create a new sheet.
                    y = 1
                    xlApp.Worksheets.add()
                    ' place field names in first row and bold them
                    For z As Integer = 0 To dgv.ColumnCount - 1
                        xlApp.Cells(y, z + 1).Value = dgv.Columns(z).HeaderText
                        xlApp.Cells(y, z + 1).Font.Bold = True
                    Next
                    y += 1
                End If

            Next

            ' loop through the worksheets to format and rename.
            ' this will give us a unique to Vision 'Sheet0'.
            sc = xlApp.WorkSheets.Count
            For i As Integer = 1 To sc
                xlApp.Worksheets(i).Columns().AutoFit()
                xlApp.Worksheets(i).Name = "Sheet" & (sc - i).ToString()
            Next

            xlApp.Visible = True
            xlApp.Quit()
            xlApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Public Shared Sub exportToWord(ByRef dgv As DataGridView)
        Dim rowCountStop As Integer = 1
        If dgv.AllowUserToAddRows = True Then
            rowCountStop = 2
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Try
            Dim stuff_to_print As String = ""
            Dim wordApp As Object = CreateObject("Word.Basic")
            wordApp.FileNew()

            ' column headings
            For i As Integer = 0 To dgv.ColumnCount - 1
                stuff_to_print &= vbTab & dgv.Columns(i).HeaderText
            Next
            wordApp.Insert(stuff_to_print)
            wordApp.InsertPara()

            ' the rest
            For i As Integer = 0 To dgv.Rows.Count - rowCountStop
                stuff_to_print = ""
                For x As Integer = 0 To dgv.ColumnCount - 1
                    stuff_to_print &= vbTab & dgv.Rows(i).Cells(x).Value.ToString()
                Next
                wordApp.Insert(stuff_to_print)
                wordApp.InsertPara()
            Next
            wordApp.AppShow()
            wordApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, My.Application.Info.Title)
        Finally
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Public Shared Function getInsertSQL(ByRef dgv As DataGridView, ByVal rowNumber As Integer) As String
        Dim sb As New System.Text.StringBuilder
        If rowNumber > -1 Then
            If dgv.DataSource.ToString.Trim.Contains(" ") = False Then
                sb.Append("INSERT INTO " & dgv.DataSource.ToString())
            Else
                sb.Append("INSERT INTO [" & dgv.DataSource.ToString() & "]")
            End If
            sb.Append(vbCrLf)
            sb.Append("(")
            For i As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Columns(i).Name.ToString.Trim.Contains(" ") = False Then
                    sb.Append(dgv.Columns(i).Name & ",")
                Else
                    sb.Append("[" & dgv.Columns(i).Name & "],")
                End If
            Next
            sb.Remove(sb.Length - 1, 1)
            sb.Append(")")
            sb.Append(vbCrLf)
            sb.Append("VALUES")
            sb.Append(vbCrLf)
            sb.Append("(")
            For i As Integer = 0 To dgv.ColumnCount - 1
                If dgv.Columns(i).ValueType.ToString = "System.String" Or dgv.Columns(i).ValueType.ToString = "System.Guid" Then
                    sb.Append("'" & VisionHelper.removeQuote(dgv.Rows(rowNumber).Cells(i).Value.ToString()) & "',")
                Else
                    If IsDBNull(dgv.Rows(rowNumber).Cells(i).Value) = False Then
                        sb.Append(dgv.Rows(rowNumber).Cells(i).Value & ",")
                    Else
                        sb.Append("Null,")
                    End If
                End If
            Next
            sb.Remove(sb.Length - 1, 1)
            sb.Append(")")
        End If
        Return sb.ToString()
    End Function

    Public Shared Function getDeleteSQL(ByRef dgv As DataGridView, ByVal rowNumber As Integer) As String
        Dim sb As New System.Text.StringBuilder
        If rowNumber > -1 Then
            If dgv.DataSource.ToString.Trim.Contains(" ") = False Then
                sb.Append("DELETE FROM " & dgv.DataSource.ToString())
            Else
                sb.Append("DELETE FROM [" & dgv.DataSource.ToString() & "]")
            End If
            sb.Append(vbCrLf)
            sb.Append("WHERE ")
            For i As Integer = 0 To dgv.ColumnCount - 1
                sb.Append(vbCrLf)
                If IsDBNull(dgv.Rows(rowNumber).Cells(i).Value) = False Then
                    If dgv.Columns(i).HeaderText.ToString.Trim.Contains(" ") = False Then
                        sb.Append(dgv.Columns(i).HeaderText & " = ")
                    Else
                        sb.Append("[" & dgv.Columns(i).HeaderText & "] = ")
                    End If
                    If dgv.Columns(i).ValueType.ToString = "System.String" Or dgv.Columns(i).ValueType.ToString = "System.Guid" Then
                        sb.Append("'" & VisionHelper.removeQuote(dgv.Rows(rowNumber).Cells(i).Value.ToString()) & "' AND ")
                    Else
                        sb.Append(dgv.Rows(rowNumber).Cells(i).Value & " AND ")
                    End If
                Else
                    If dgv.Columns(i).HeaderText.ToString.Trim.Contains(" ") = False Then
                        sb.Append(dgv.Columns(i).HeaderText & " is Null AND ")
                    Else
                        sb.Append("[" & dgv.Columns(i).HeaderText & "] is Null AND ")
                    End If
                End If
            Next
            sb.Remove(sb.Length - 4, 3)
        End If
        Return sb.ToString()
    End Function

    Public Shared Function getUpdateSQL(ByRef dgv As DataGridView, ByVal rowNumber As Integer, ByVal originalValues As DataGridViewRow) As String
        Dim sb As New System.Text.StringBuilder
        If rowNumber > -1 Then
            If dgv.DataSource.ToString.Trim.Contains(" ") = False Then
                sb.Append("UPDATE " & dgv.DataSource.ToString())
            Else
                sb.Append("UPDATE [" & dgv.DataSource.ToString() & "]")
            End If
            sb.Append(vbCrLf)
            sb.Append("SET ")
            For i As Integer = 0 To dgv.ColumnCount - 1
                sb.Append(vbCrLf)
                If dgv.Columns(i).HeaderText.ToString.Trim.Contains(" ") = False Then
                    sb.Append(dgv.Columns(i).HeaderText & " = ")
                Else
                    sb.Append("[" & dgv.Columns(i).HeaderText & "] = ")
                End If
                If IsDBNull(dgv.Rows(rowNumber).Cells(i).Value) = False Then
                    If dgv.Columns(i).ValueType.ToString = "System.String" Or dgv.Columns(i).ValueType.ToString = "System.Guid" Then
                        sb.Append("'" & VisionHelper.removeQuote(dgv.Rows(rowNumber).Cells(i).Value.ToString()) & "', ")
                    Else
                        sb.Append(dgv.Rows(rowNumber).Cells(i).Value & ", ")
                    End If
                Else
                    sb.Append("Null, ")
                End If
            Next
            sb.Remove(sb.Length - 2, 2)
            sb.Append(vbCrLf)
            sb.Append("WHERE ")
            For i As Integer = 0 To originalValues.Cells.Count - 1
                sb.Append(vbCrLf)
                If IsDBNull(originalValues.Cells(i).Value) = False Then
                    If dgv.Columns(i).HeaderText.ToString.Trim.Contains(" ") = False Then
                        sb.Append(dgv.Columns(i).HeaderText & " = ")
                    Else
                        sb.Append("[" & dgv.Columns(i).HeaderText & "] = ")
                    End If
                    If originalValues.Cells(i).ValueType.FullName = "System.String" Or originalValues.Cells(i).ValueType.FullName = "System.Guid" Then
                        sb.Append("'" & VisionHelper.removeQuote(originalValues.Cells(i).Value.ToString()) & "' AND ")
                    Else
                        sb.Append(originalValues.Cells(i).Value & " AND ")
                    End If
                Else
                    If dgv.Columns(i).HeaderText.ToString.Trim.Contains(" ") = False Then
                        sb.Append(dgv.Columns(i).HeaderText & " is Null AND ")
                    Else
                        sb.Append("[" & dgv.Columns(i).HeaderText & "] is Null AND ")
                    End If
                End If
            Next
            sb.Remove(sb.Length - 4, 3)
        End If
        Return sb.ToString()
    End Function

    Public Shared Function removeQuote(ByVal value As String)
        If value.Trim.Length > 0 Then
            Return value.Replace("'", "''")
        Else
            Return value
        End If
    End Function

    Public Shared Sub getDataSourceNames(ByRef Combo1 As System.Windows.Forms.ComboBox)
        'Get all ODBC Datasource Names.
        '==============================
        Dim bSetDefault As Boolean
        Dim SQL_SUCCESS As Integer = 0
        Dim SQL_FETCH_NEXT As Integer = 1

        'Get the current default DSN settings from the registry.
        Dim DefaultDSN As String = GetSetting(My.Application.Info.Title, "Settings", "DefaultDSN", "")
        bSetDefault = Len(Trim(DefaultDSN))

        On Error Resume Next

        Dim i As Short
        Dim sDSNItem As New String("", 1024)
        Dim sDRVItem As New String("", 1024)
        Dim sDSN As String
        Dim sDRV As String
        Dim iDSNLen As Short
        Dim iDRVLen As Short
        Dim lHenv As Integer 'handle to the environment

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Get the DSNs.
        '=============
        If SQLAllocEnv(lHenv) <> -1 Then
            Combo1.Items.Clear()
            Combo1.Sorted = True
            Combo1.Items.Add("")
            Do Until i <> SQL_SUCCESS
                i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
                If i = SQL_SUCCESS Then
                    sDSN = Left(sDSNItem, iDSNLen)
                    sDRV = Left(sDRVItem, iDRVLen)
                    If sDSN <> Space(iDSNLen) Then
                        Combo1.Items.Add(sDSN)
                    End If
                End If
            Loop
            If Not bSetDefault Then
                ' set our list to the first item.
                Combo1.SelectedIndex = 0
            Else
                Combo1.SelectedIndex = Combo1.FindString(DefaultDSN)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Public Shared Sub getDataSourceNames(ByRef list1 As System.Windows.Forms.ListView)
        'Get all ODBC Datasource Names.
        '==============================
        Dim bSetDefault As Boolean
        Dim SQL_SUCCESS As Integer = 0
        Dim SQL_FETCH_NEXT As Integer = 1

        'Get the current default DSN settings from the registry.
        Dim DefaultDSN As String = GetSetting(My.Application.Info.Title, "Settings", "DefaultDSN", "")
        bSetDefault = Len(Trim(DefaultDSN))

        On Error Resume Next

        Dim i As Short
        Dim sDSNItem As New String("", 1024)
        Dim sDRVItem As New String("", 1024)
        Dim sDSN As String
        Dim sDRV As String
        Dim iDSNLen As Short
        Dim iDRVLen As Short
        Dim lHenv As Integer 'handle to the environment

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'Get the DSNs.
        '=============
        If SQLAllocEnv(lHenv) <> -1 Then
            list1.Items.Clear()
            Do Until i <> SQL_SUCCESS
                i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
                If i = SQL_SUCCESS Then
                    sDSN = Left(sDSNItem, iDSNLen)
                    sDRV = Left(sDRVItem, iDRVLen)
                    If sDSN <> Space(iDSNLen) Then
                        Dim itm As New ListViewItem(sDSN)
                        Dim vu As New VisionUser(sDSN)
                        itm.SubItems.Add(vu.UserID)
                        list1.Items.Add(itm)
                    End If
                End If
            Loop
            If Not bSetDefault Then
                ' set our list to the first item.
                list1.Items(0).Selected = True
            Else
                list1.FindItemWithText(DefaultDSN).Selected = True
                list1.FocusedItem = list1.FindItemWithText(DefaultDSN)
            End If
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Public Shared Sub callODBC()
        'Call the ODBC Driver Manager.
        '=============================
        Try
            Dim lpBuffer As New String("", 255)
            Dim sysPath As String
            Dim sysDir As Integer

            sysDir = apiGetSystemDirectory(lpBuffer, Len(lpBuffer))
            sysPath = Left(lpBuffer, sysDir)

            If sysPath <> "" Then
                'Spawns the ODBC Admin.
                '======================
                Shell(sysPath & "\ODBCAD32.exe", AppWinStyle.NormalFocus)
            End If
        Catch ex As Exception
            MsgBox("Unable to locate ODBCAD32.EXE.  Please use Control Panel to access ODBC Administrator.", MsgBoxStyle.Exclamation, My.Application.Info.Title)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Public Shared Function Encrypt(ByVal s) As String
        Dim rtnValue As String = s
        For x As Integer = 1 To Len(s)
            Mid(rtnValue, x, 1) = Chr(Asc(Mid(rtnValue, x, 1)) Xor 6)
        Next
        Return rtnValue
    End Function

End Class
