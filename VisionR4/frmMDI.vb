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
Public Class mdiForm

    Private Sub mdiForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Me.WindowState = FormWindowState.Minimized Then
            ' do nothing :-)
        Else
            My.Settings.WindowState = Me.WindowState
            My.Settings.ClientSize = Me.ClientSize
            My.Settings.Save()
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ShowSplashToolStripMenuItem.Checked = CType(GetSetting(My.Application.Info.Title, "Settings", "ShowSplash", "True"), Boolean)
        Me.PlayAboutBoxSoundToolStripMenuItem.Checked = CType(GetSetting(My.Application.Info.Title, "Settings", "PlaySound", "True"), Boolean)
        If Me.ShowSplashToolStripMenuItem.Checked = True Then
            Dim sc As New SplashScreen
            sc.ShowDialog()
        End If

        Me.WindowState = My.Settings.WindowState
        Me.ClientSize = My.Settings.ClientSize
    End Sub

    Private Sub ExitApplication()
        Me.Close()
    End Sub

    Private Sub showNewQuery()
        Dim f As New frmNewQuery
        f.MdiParent = Me
        f.Show()
    End Sub

    Private Sub openSavedQuery()
        Dim f As New frmSavedQueries
        f.MdiParent = Me
        f.Show()
    End Sub

    Private Sub setUserIDandPassword()
        Dim f As New frmUserIds
        f.ShowDialog()
    End Sub

    Private Sub setCommandTimeOut()
        Dim f As New frmCmdTimeOut
        f.ShowDialog()
    End Sub

    Private Sub setDefaultDSN()
        Dim f As New frmDSNDefault
        f.ShowDialog()
    End Sub

    Private Sub showUserSQL()
        Dim frmUpdateSQL As New frmUserSQL
        frmUpdateSQL.MdiParent = Me
        frmUpdateSQL.Show()
    End Sub

    Private Sub showAbout()
        Dim ab As New frmAboutBox
        ab.playCrashSound = Me.PlayAboutBoxSoundToolStripMenuItem.Checked
        ab.ShowDialog()
    End Sub

    Private Sub showProcessScript()
        Dim fs As New frmScript
        fs.MdiParent = Me
        fs.Show()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Me.ExitApplication()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.ExitApplication()
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Me.showNewQuery()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Me.showAbout()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Me.openSavedQuery()
    End Sub

    Private Sub CustomizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomizeToolStripMenuItem.Click
        Me.setCommandTimeOut()
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Me.setUserIDandPassword()
    End Sub

    Private Sub DataSourceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataSourceToolStripMenuItem.Click
        Me.setDefaultDSN()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.showNewQuery()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.openSavedQuery()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        VisionHelper.callODBC()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.setUserIDandPassword()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Me.setCommandTimeOut()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Me.setDefaultDSN()
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Me.showAbout()
    End Sub

    Private Sub UserSQLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserSQLToolStripMenuItem.Click
        Me.showUserSQL()
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        Me.showUserSQL()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseAllToolStripMenuItem.Click
        For Each child As Form In Me.MdiChildren
            child.Close()
        Next
    End Sub

    Private Sub UndoToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        VisionHelper.callODBC()
    End Sub

    Private Sub ShowSplashToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowSplashToolStripMenuItem.Click
        SaveSetting(My.Application.Info.Title, "Settings", "ShowSplash", Me.ShowSplashToolStripMenuItem.Checked.ToString)
    End Sub

    Private Sub PlayAboutBoxSoundToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlayAboutBoxSoundToolStripMenuItem.Click
        SaveSetting(My.Application.Info.Title, "Settings", "PlaySound", Me.PlayAboutBoxSoundToolStripMenuItem.Checked.ToString)
    End Sub

    Private Sub RunScriptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunScriptToolStripMenuItem.Click
        Me.showProcessScript()
    End Sub

    Private Sub ToolStripButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton10.Click
        Me.showProcessScript()
    End Sub
End Class
