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
Public Class frmDSNDefault

    Private Sub frmDSNDefault_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        VisionHelper.getDataSourceNames(Me.cboDSNs)
    End Sub

    Private Sub exitApplication()
        Me.Close()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.exitApplication()
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        If IsNothing(Me.cboDSNs.SelectedItem) = False Then
            SaveSetting(My.Application.Info.Title, "Settings", "DefaultDSN", Me.cboDSNs.SelectedItem)
        Else
            SaveSetting(My.Application.Info.Title, "Settings", "DefaultDSN", "")
        End If
        Me.exitApplication()
    End Sub

End Class