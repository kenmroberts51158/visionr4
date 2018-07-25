Public Class frmUserIds

    Private Sub frmUserIds_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VisionHelper.getDataSourceNames(Me.lstDsns)

    End Sub

    Private Sub btnDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefault.Click
        Dim idx As Integer = lstDsns.SelectedItems(0).Index
        Dim f As New frmUSR
        f.ShowDialog()
        VisionHelper.getDataSourceNames(Me.lstDsns)
        'Re-select the item that was selected before rebuilding the list.
        lstDsns.Items(idx).Selected = True
        lstDsns.FocusedItem = lstDsns.Items(idx)
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModify.Click
        Dim idx As Integer = lstDsns.SelectedItems(0).Index
        Dim f As New frmUSR(lstDsns.SelectedItems(0).SubItems(0).Text)
        f.ShowDialog()
        VisionHelper.getDataSourceNames(Me.lstDsns)
        'Re-select the item that was selected before rebuilding the list.
        lstDsns.Items(idx).Selected = True
        lstDsns.FocusedItem = lstDsns.Items(idx)
    End Sub

    Private Sub lstDsns_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstDsns.DoubleClick
        btnModify.PerformClick()
    End Sub

    Private Sub lstDsns_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstDsns.KeyUp
        If e.KeyCode = Keys.Enter Then
            btnModify.PerformClick()
        End If
    End Sub
End Class