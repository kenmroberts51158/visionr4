
Public Class ColumnSchemaManager

    Public Shared Function GetColumnList() As ColumnSchema()
        Dim rtnVal() As ColumnSchema

        'open xml file
        Dim folder As String
        ''folder = System.Environment.CurrentDirectory
        folder = Application.StartupPath

        Dim dataSet1 As New DataSet()
        dataSet1.ReadXml(folder & "\ColumnsSchema.xml")
        ReDim rtnVal(dataSet1.Tables(0).Rows.Count - 1)

        Dim rw As Data.DataRow
        Dim i As Integer = 0
        For Each rw In dataSet1.Tables(0).Rows
            rtnVal(i) = New ColumnSchema(rw)
            i += 1
        Next

        Return rtnVal
    End Function
End Class

Public Class ColumnSchema
    Private c_Row As Data.DataRow

    Public Sub New(ByRef rw As Data.DataRow)
        c_Row = rw
    End Sub

    Public ReadOnly Property ColumnName() As String
        Get
            Return c_Row("name")
        End Get
    End Property
End Class
