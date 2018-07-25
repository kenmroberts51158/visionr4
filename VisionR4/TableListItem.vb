''' <summary>
''' Class used to create entries in the Table Combo box that has the description concatenated to the table name.
''' </summary>
''' <remarks></remarks>
Public Class TableListItem
    Public Property TableName As String
    Public Property TableDescription As String

    Public Sub New(Name As String, Description As String)
        TableName = Name.Trim()
        TableDescription = Description.Trim()
    End Sub

    Public Overrides Function ToString() As String
        Dim rtnVal As String
        If TableDescription.Length > 0 Then
            rtnVal = TableName & "  -  " & TableDescription
        Else
            rtnVal = TableName
        End If
        Return rtnVal
    End Function

    Private Sub New()
        'to keep it from being accessible.
    End Sub

End Class
