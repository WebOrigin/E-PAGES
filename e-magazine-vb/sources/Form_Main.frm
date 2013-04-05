VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Begin VB.Form WOEM_Main 
   Caption         =   "WebOrigin"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   12780
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_Ac 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8775
      Top             =   6795
   End
   Begin VB.PictureBox AImg_F 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   540
      ScaleHeight     =   5910
      ScaleWidth      =   6405
      TabIndex        =   5
      Top             =   1575
      Visible         =   0   'False
      Width           =   6405
      Begin VB.PictureBox AImg_Rol_F_F 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   180
         ScaleHeight     =   960
         ScaleWidth      =   6000
         TabIndex        =   6
         Top             =   4815
         Width           =   6000
         Begin VB.PictureBox AImg_Rol_F 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   315
            ScaleHeight     =   870
            ScaleWidth      =   3705
            TabIndex        =   7
            Top             =   180
            Width           =   3705
            Begin VB.PictureBox AImg_Rol 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1140
               Left            =   450
               ScaleHeight     =   1140
               ScaleWidth      =   6720
               TabIndex        =   8
               Top             =   90
               Width           =   6720
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   19
                  Left            =   90
                  MouseIcon       =   "Form_Main.frx":0000
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   18
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0152
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   17
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":02A4
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   16
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":03F6
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   15
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0548
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   14
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":069A
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   13
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":07EC
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   12
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":093E
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   11
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0A90
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   10
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0BE2
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   9
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0D34
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   8
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0E86
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   7
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":0FD8
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   6
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":112A
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   5
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":127C
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   4
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":13CE
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   3
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":1520
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   2
                  Left            =   0
                  MouseIcon       =   "Form_Main.frx":1672
                  MousePointer    =   99  'Custom
                  Top             =   0
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   1
                  Left            =   765
                  MouseIcon       =   "Form_Main.frx":17C4
                  MousePointer    =   99  'Custom
                  Top             =   270
                  Width           =   1230
               End
               Begin VB.Image AImg_Tmb 
                  Height          =   780
                  Index           =   0
                  Left            =   225
                  MouseIcon       =   "Form_Main.frx":1916
                  MousePointer    =   99  'Custom
                  Top             =   180
                  Width           =   1230
               End
            End
         End
      End
      Begin VB.Image AImg_Show 
         Height          =   3165
         Left            =   180
         Top             =   180
         Width           =   6045
      End
   End
   Begin VB.PictureBox Picture_bt_b 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   5895
      MouseIcon       =   "Form_Main.frx":1A68
      MousePointer    =   99  'Custom
      Picture         =   "Form_Main.frx":1D72
      ScaleHeight     =   630
      ScaleWidth      =   3690
      TabIndex        =   4
      Top             =   135
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.PictureBox Picture_bt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2520
      MouseIcon       =   "Form_Main.frx":7D2D
      MousePointer    =   99  'Custom
      Picture         =   "Form_Main.frx":8037
      ScaleHeight     =   510
      ScaleWidth      =   3045
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   3045
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Lsk_Flash 
      Height          =   700
      Index           =   2
      Left            =   1610
      TabIndex        =   1
      Top             =   70
      Width           =   700
      _cx             =   1235
      _cy             =   1235
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Lsk_Flash 
      Height          =   700
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   70
      Width           =   700
      _cx             =   1235
      _cy             =   1235
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Lsk_Flash 
      Height          =   700
      Index           =   0
      Left            =   70
      TabIndex        =   2
      Top             =   70
      Width           =   700
      _cx             =   1235
      _cy             =   1235
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "WOEM_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fod_ID_St As Integer
Dim AI_DX As Single
Dim AI_DLeft As Single
Dim Time_Way As Integer

Private Sub AImg_F_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer_Ac.Enabled = True
End Sub

Private Sub AImg_Rol_F_F_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer_Ac.Enabled = True
End Sub

Private Sub AImg_Rol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
AI_DX = X
AI_DLeft = AImg_Rol.Left
End Sub

Private Sub AImg_Rol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer_Ac.Enabled = False
If Button <> 1 Then Exit Sub
Dim AI_New_X As Single
AI_New_X = AImg_Rol.Left + (X - AI_DX)
If AI_New_X > 0 Then AI_New_X = 0
If AI_New_X + AImg_Rol.Width < AImg_Rol_F.Width Then AI_New_X = AImg_Rol_F.Width - AImg_Rol.Width
AImg_Rol.Left = AI_New_X
End Sub



Private Sub AImg_Rol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If AImg_Rol.Left > AI_DLeft Then Time_Way = -1
If AImg_Rol.Left < AI_DLeft Then Time_Way = 1
End Sub

Private Sub AImg_Show_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer_Ac.Enabled = True
End Sub

Private Sub AImg_Tmb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
AImg_Show.Picture = LoadPicture(App.Path & "\AIMG_" & Fod_ID_St & "\" & Index & ".jpg")
AI_DX = X
AI_DLeft = AImg_Rol.Left
End Sub

Private Sub AImg_Tmb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer_Ac.Enabled = False
If Button <> 1 Then Exit Sub
Dim AI_New_X As Single
AI_New_X = AImg_Rol.Left + (X - AI_DX)
If AI_New_X > 0 Then AI_New_X = 0
If AI_New_X + AImg_Rol.Width < AImg_Rol_F.Width Then AI_New_X = AImg_Rol_F.Width - AImg_Rol.Width
AImg_Rol.Left = AI_New_X
End Sub

Private Sub AImg_Tmb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If AImg_Rol.Left > AI_DLeft Then Time_Way = -1
If AImg_Rol.Left < AI_DLeft Then Time_Way = 1
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
Lsk_Flash(0).Movie = App.Path + "\Magazine.swf"
Lsk_Flash(0).Play
Lsk_Flash(1).Visible = False
Lsk_Flash(2).Visible = False
'Load_AIMG 1, 10
End Sub

Private Sub Form_Resize()
On Error Resume Next
Lsk_Flash(0).Left = 0
Lsk_Flash(0).Top = 0
Lsk_Flash(0).Width = Me.Width
Lsk_Flash(0).Height = Me.Height - 547
End Sub

Private Sub Lsk_Flash_FSCommand(Index As Integer, ByVal command As String, ByVal args As String)
'Me.Caption = command
Select Case LCase(command)
    Case "page_now,2,3"
        Lsk_Flash(1).Visible = True
        Lsk_Flash(1).Width = 6000
        Lsk_Flash(1).Height = 4500
        Lsk_Flash(1).Top = Int((Me.Height - Lsk_Flash(1).Height) / 2) + 1300
        Lsk_Flash(1).Left = Int((Me.Width - Lsk_Flash(1).Width) / 2) + 3800
        Lsk_Flash(1).Movie = App.Path + "\3D_1.swf"
        Lsk_Flash(1).Play
    Case "page_now,4,5"
        Lsk_Flash(1).Visible = True
        Lsk_Flash(1).Width = 320 * 15
        Lsk_Flash(1).Height = 240 * 15
        Lsk_Flash(1).Top = Int((Me.Height - Lsk_Flash(1).Height) / 2) + 500
        Lsk_Flash(1).Left = Int((Me.Width - Lsk_Flash(1).Width) / 2) - 4000
        Lsk_Flash(1).Movie = App.Path + "\video.swf"
        Lsk_Flash(1).Play
        Load_AIMG 1, 10
        AImg_F.Top = Int((Me.Height - Lsk_Flash(1).Height) / 2) - 400
        AImg_F.Left = Int((Me.Width - Lsk_Flash(1).Width) / 2) + 3000
        AImg_F.Visible = True
    Case "page_now,6,7"
        Picture_bt.Visible = True
        Picture_bt.Top = Int((Me.Height - Lsk_Flash(1).Height) / 2) + 4000
        Picture_bt.Left = Int((Me.Width - Lsk_Flash(1).Width) / 2) - 3000
    Case "page_now,10,11"
        Picture_bt_b.Visible = True
        Picture_bt_b.Top = Int((Me.Height - Lsk_Flash(1).Height) / 2) - 315
        Picture_bt_b.Left = Int((Me.Width - Lsk_Flash(1).Width) / 2) - 3330
    Case "screen_full"
        Lsk_Flash(1).Visible = False
        Select Case Me.WindowState
            Case 0
                Me.WindowState = 2
            Case 2
                Me.WindowState = 0
        End Select
End Select

If Left(command, 8) = "page_now" Then
    If command <> "page_now,2,3" And command <> "page_now,4,5" Then
        Lsk_Flash(1).Visible = False
        Lsk_Flash(1).Movie = App.Path + "\0.swf"
    End If
End If

If command <> "page_now,6,7" Then
    Picture_bt.Visible = False
End If
If command <> "page_now,10,11" Then
    Picture_bt_b.Visible = False
End If
If command <> "page_now,4,5" Then
    AImg_F.Visible = False
    Timer_Ac.Enabled = False
End If
End Sub

Private Sub Picture_bt_b_Click()
Shell "explorer.exe http://www.weborigin.co.nz/"
End Sub

Private Sub Picture_bt_Click()
Shell "explorer.exe " & App.Path & "\paypal.pdf"
End Sub


Private Sub Load_AIMG(Fod_ID As Integer, Pic_Num As Integer)
Fod_ID_St = Fod_ID
Dim i As Integer
Dim PNum As Integer
For i = 0 To 19
    AImg_Tmb(i).Visible = False
    AImg_Tmb(i).Top = 0
Next
AImg_Tmb(0).Left = 0
For i = 0 To 10
    AImg_Tmb(i).Width = 67 * 15
    AImg_Tmb(i).Height = 50 * 75
    AImg_Tmb(i).Picture = LoadPicture(App.Path & "\AIMG_" & Fod_ID & "\" & i & "_s.jpg")
    AImg_Tmb(i).Visible = True
    If i > 0 Then AImg_Tmb(i).Left = AImg_Tmb(i - 1).Left + 67 * 15 + 70
Next
AImg_Show.Picture = LoadPicture(App.Path & "\AIMG_" & Fod_ID & "\1.jpg")
AImg_F.Picture = LoadPicture(App.Path & "\AIMG_" & Fod_ID & "\ai_bg.jpg")
AImg_Rol.Width = AImg_Tmb(10).Left + AImg_Tmb(10).Width
AImg_Rol.Height = 50 * 15
AImg_Rol_F.Height = 50 * 15
AImg_Rol_F.Width = AImg_Rol_F_F.Width - 200
AImg_Rol_F.Left = (AImg_Rol_F_F.Width - AImg_Rol_F.Width) / 2
AImg_Rol_F.Top = (AImg_Rol_F_F.Height - AImg_Rol_F.Height) / 2
AImg_Rol.BackColor = &HE0E0E0    'RGB(51, 51, 51)
AImg_Rol_F.BackColor = &HE0E0E0   ' RGB(51, 51, 51)
AImg_Rol.Left = 0
AImg_Rol.Top = 0
Time_Way = 1
Timer_Ac.Enabled = True
End Sub

Private Sub Timer_Ac_Timer()
Select Case Time_Way
    Case 1
        AImg_Rol.Left = AImg_Rol.Left - 30
        If AImg_Rol.Left + AImg_Rol.Width < AImg_Rol_F.Width Then Time_Way = -1
    Case -1
        AImg_Rol.Left = AImg_Rol.Left + 30
        If AImg_Rol.Left > 0 Then Time_Way = 1
End Select
End Sub
