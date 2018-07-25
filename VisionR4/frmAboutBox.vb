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
Public Class frmAboutBox
    Private Declare Auto Function PlaySound Lib "winmm.dll" (ByVal pszSound As Byte(), ByVal hModule As IntPtr, ByVal dwFlags As Int32) As Boolean
    Public playCrashSound As Boolean = True

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For i As Decimal = 1.0 To 0.1 Step -0.1
            Me.Opacity = i
            Me.Refresh()
            System.Threading.Thread.Sleep(100)
        Next
        Me.Close()
    End Sub

    Private Sub frmAboutBox_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.tryToStopCrashSound()
    End Sub

    Private Sub frmAboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Label2.Text = "Vision R4: " & My.Application.Info.Version.ToString
    End Sub

    Public Sub crash()
        If playCrashSound = False Then Exit Sub

        Try
            Dim playAsync As Int32 = &H1
            Dim playFromStream As Int32 = &H4
            Dim soundName As String = "Kamikaze.wav"
            Dim appName As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
            Dim soundStream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(appName & "." & soundName)
            Dim soundStreamLength As Integer = CInt(soundStream.Length)
            Dim soundByteArray(soundStreamLength - 1) As Byte
            soundStream.Read(soundByteArray, IntPtr.Zero, soundStreamLength)
            soundStream.Close()
            PlaySound(soundByteArray, IntPtr.Zero, playFromStream Or playAsync)
        Catch ex As Exception
            ' eaten
        End Try
    End Sub

    Private Sub tryToStopCrashSound()
        Me.BackgroundWorker1.CancelAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.crash()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.BackgroundWorker1.RunWorkerAsync()
        Me.Timer1.Enabled = False
    End Sub
End Class


