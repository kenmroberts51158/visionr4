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
Public Class frmUSR

    Private visionUser As New VisionUser()

    Private myDSN As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal dsn As String)
        InitializeComponent()

        Me.visionUser.getDefaults(dsn)
        myDSN = dsn

        Me.Text = dsn & " User ID & Password"

    End Sub

    Private Sub exitAppliation()
        Me.Close()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.exitAppliation()
    End Sub

    Private Sub frmUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub

    Private Sub frmUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.visionUser = Nothing
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        If myDSN = "" Then
            Me.visionUser.setDefaults(Me.txtUserId.Text, Me.txtPassword.Text)
        Else
            Me.visionUser.setDefaults(myDSN, Me.txtUserId.Text, Me.txtPassword.Text)
        End If
        Me.exitAppliation()
    End Sub

    Private Sub btmRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btmRestore.Click
        If myDSN = "" Then
            Me.visionUser.restoreDefaults()
        Else
            ''Me.visionUser.restoreDefaults(dsn)
        End If

        Me.txtUserId.Text = Me.visionUser.UserID
        Me.txtPassword.Text = Me.visionUser.Password
    End Sub
End Class