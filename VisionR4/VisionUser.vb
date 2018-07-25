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
Public Class VisionUser

    Private userName As String = ""
    Private userPassword As String = ""

    Public Property UserID() As String
        Get
            Return Me.userName
        End Get
        Set(ByVal value As String)
            Me.userName = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return Me.userPassword
        End Get
        Set(ByVal value As String)
            Me.userPassword = value
        End Set
    End Property

    Public Sub New()
        Me.getDefaults()
    End Sub

    Public Sub New(ByVal userID As String, ByVal password As String)
        Me.userName = userID
        Me.userPassword = password
    End Sub

    Public Sub New(ByVal dsn As String)
        getDefaults(dsn)

    End Sub

    Private Sub getDefaults()
        'Get the current settings from the registry.
        Me.userName = GetSetting(My.Application.Info.Title, "Settings", "UID", "")
        Me.userPassword = VisionHelper.Encrypt(GetSetting(My.Application.Info.Title, "Settings", "PWD", ""))
    End Sub

    Public Sub getDefaults(ByVal dsn As String)
        'Get the current settings from the registry.
        Me.userName = GetSetting(My.Application.Info.Title, "Settings\IDS\" & dsn, "UID", "")
        Me.userPassword = VisionHelper.Encrypt(GetSetting(My.Application.Info.Title, "Settings\IDS\" & dsn, "PWD", ""))

        If Me.userName = "" And Me.userPassword = "" Then
            getDefaults()
        End If
    End Sub

    Public Sub setDefaults(ByVal userID As String, ByVal password As String)
        'Save the current settings to the registry.
        SaveSetting(My.Application.Info.Title, "Settings", "UID", userID)
        SaveSetting(My.Application.Info.Title, "Settings", "PWD", VisionHelper.Encrypt(password))
    End Sub

    Public Sub setDefaults(ByVal dsn As String, ByVal userID As String, ByVal password As String)
        'Save the current settings to the registry.
        SaveSetting(My.Application.Info.Title, "Settings\IDS\" & dsn, "UID", userID)
        SaveSetting(My.Application.Info.Title, "Settings\IDS\" & dsn, "PWD", VisionHelper.Encrypt(password))
    End Sub

    Public Sub restoreDefaults()
        Me.getDefaults()
    End Sub

End Class
