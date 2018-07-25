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
Public Class frmCmdTimeOut

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub frmCmdTimeOut_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtCmdTime.Text = VisionHelper.CommandTimeout
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        VisionHelper.restoreCommandTimeoutToDefault()
        Me.txtCmdTime.Text = VisionHelper.CommandTimeout
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        If IsNumeric(Me.txtCmdTime.Text) Then
            Me.txtCmdTime.Text = CType(Me.txtCmdTime.Text, Integer).ToString
            VisionHelper.CommandTimeout = Me.txtCmdTime.Text
            Me.Close()
        Else
            MsgBox("Command Time-out value must be a number.", MsgBoxStyle.Exclamation, My.Application.Info.Title)
        End If
    End Sub

End Class