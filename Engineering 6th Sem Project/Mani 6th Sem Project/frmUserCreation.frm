VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmUserCreation 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Craeting User Account - *"
   ClientHeight    =   4545
   ClientLeft      =   1740
   ClientTop       =   1995
   ClientWidth     =   8205
   Icon            =   "frmUserCreation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4545
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   270
      Left            =   1395
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   476
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3225
      Left            =   285
      TabIndex        =   7
      Top             =   885
      Visible         =   0   'False
      Width           =   7725
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Book Dealer Database"
         Enabled         =   0   'False
         Height          =   345
         Index           =   5
         Left            =   3180
         TabIndex        =   21
         Tag             =   " "
         Top             =   720
         Width           =   2640
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Deleting Temporary Files"
         Enabled         =   0   'False
         Height          =   345
         Index           =   9
         Left            =   3180
         TabIndex        =   20
         Tag             =   " "
         Top             =   2205
         Width           =   2415
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Student Enrollment Database"
         Enabled         =   0   'False
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   19
         Tag             =   " "
         Top             =   2205
         Width           =   3150
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Order Processing Database"
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Tag             =   " "
         Top             =   1833
         Width           =   3120
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Insurance Database"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   75
         TabIndex        =   17
         Tag             =   " "
         Top             =   1462
         Width           =   2910
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Oracle DataBase User"
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   75
         TabIndex        =   16
         Tag             =   " "
         Top             =   1110
         Width           =   2715
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saving Theme Settings"
         Enabled         =   0   'False
         Height          =   345
         Index           =   8
         Left            =   3165
         TabIndex        =   15
         Tag             =   " "
         Top             =   1830
         Width           =   2415
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opening Connection"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Tag             =   " "
         Top             =   735
         Width           =   2160
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saving  User Information"
         Enabled         =   0   'False
         Height          =   345
         Index           =   7
         Left            =   3180
         TabIndex        =   13
         Tag             =   " "
         Top             =   1455
         Width           =   2145
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creating Bank Database"
         Enabled         =   0   'False
         Height          =   345
         Index           =   6
         Left            =   3180
         TabIndex        =   12
         Tag             =   " "
         Top             =   1095
         Width           =   2280
      End
      Begin ManoharButton.MyButton cmdShowDetail 
         Height          =   405
         Left            =   30
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "Show Details Of All Above"
         Top             =   2745
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "S&how Details"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUserCreation.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton cmdOkContinue 
         Height          =   405
         Left            =   2617
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "Continue Furhter"
         Top             =   2745
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "O&k Continue"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUserCreation.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton cmdExitNow 
         Height          =   405
         Left            =   5220
         TabIndex        =   27
         Tag             =   "1"
         ToolTipText     =   "Exit Application"
         Top             =   2745
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "E&xit Now"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUserCreation.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblHeading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configuring System Registry, DSN && Database"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   390
         Left            =   105
         TabIndex        =   8
         Tag             =   " "
         Top             =   120
         Width           =   7200
      End
      Begin VB.Image imgPhoto 
         Height          =   1710
         Left            =   5820
         Tag             =   "1"
         Top             =   810
         Width           =   1740
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   270
      TabIndex        =   22
      Top             =   900
      Width           =   7740
      Begin VB.ComboBox cmbThemes 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   255
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Select a theme"
         Top             =   1005
         Width           =   1500
      End
      Begin XPTEXTBOX.text txtUserName 
         Height          =   390
         Left            =   4545
         TabIndex        =   0
         Tag             =   "1"
         ToolTipText     =   "Enter Login Name"
         Top             =   330
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   "mani"
         BackColor       =   4194304
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ManoharButton.MyButton cmdCreate 
         Height          =   405
         Left            =   4545
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Create User"
         Top             =   2430
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "C&reate"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUserCreation.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton cmExit 
         Height          =   405
         Left            =   6180
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Exit Application"
         Top             =   2445
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUserCreation.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin XPTEXTBOX.text txtPassword 
         Height          =   390
         Left            =   4575
         TabIndex        =   1
         Tag             =   "2"
         ToolTipText     =   "Enter Password"
         Top             =   975
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "Webdings"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   "mani"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPTEXTBOX.text txtVerify 
         Height          =   390
         Left            =   4545
         TabIndex        =   2
         Tag             =   "2"
         ToolTipText     =   "Enter Varification Password"
         Top             =   1635
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "Webdings"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   "mani"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgTheme 
         Height          =   1470
         Left            =   255
         Top             =   1605
         Width           =   1500
      End
      Begin VB.Image imgAccept 
         Height          =   435
         Left            =   2250
         Top             =   2460
         Width           =   375
      End
      Begin VB.Label lblAccepted 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Accepted"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2790
         TabIndex        =   28
         Tag             =   "1"
         Top             =   2505
         Width           =   1665
      End
      Begin VB.Label lblThemeSelect 
         BackStyle       =   0  'Transparent
         Caption         =   "Theme Setting"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   270
         TabIndex        =   26
         Tag             =   "1"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblVerify 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   2295
         TabIndex        =   25
         Tag             =   "1"
         Top             =   1665
         Width           =   2385
      End
      Begin VB.Label lblUname 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   2295
         TabIndex        =   24
         Tag             =   "1"
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   2295
         TabIndex        =   23
         Tag             =   "1"
         Top             =   1005
         Width           =   2385
      End
   End
   Begin VB.Image imgRestore 
      Height          =   270
      Left            =   7515
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   105
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   165
      Stretch         =   -1  'True
      ToolTipText     =   "Application Icon"
      Top             =   75
      Width           =   315
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   7200
      Stretch         =   -1  'True
      ToolTipText     =   "Minimise"
      Top             =   105
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   7830
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   105
      Width           =   285
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5370
      Top             =   15
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creating User Account - *"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   975
      TabIndex        =   6
      Tag             =   "1"
      Top             =   60
      Width           =   2580
   End
End
Attribute VB_Name = "frmUserCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Dim errStr, errDes As String
Dim ThemeIndex As Integer
Dim accepted As Boolean
Private Sub cmbThemes_Click()
On Error GoTo call_err_handle
Call Change_Theme(cmbThemes.Text)
Call Apply_Theme(Me, 1)
Call Load_PasswordChar(txtPassword)
Call Load_PasswordChar(txtVerify)

If accepted Then
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\accepted.jpg")
lblAccepted.Caption = "Accepted"
Else
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\notaccepted.jpg")
lblAccepted.Caption = "Not Accepted"
End If

call_err_handle:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "Information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If

End Sub

Private Sub cmbThemes_GotFocus()
Call cc.Speak("Select a theme to use.")
End Sub

Private Sub cmdCreate_Click()
Call click_ok
End Sub

Private Sub cmdExitNow_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
End Sub

Private Sub cmdOkContinue_Click()
Unload Me
frmLogin.Show
End Sub
Private Sub cmdShowDetail_Click()
frmTextEditor.rtfFile.Locked = True
frmTextEditor.Show
End Sub
Private Sub cmExit_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
End Sub
Private Sub Form_Activate()
Call cmbThemes_Click

End Sub

Private Sub Form_DblClick()
Me.Left = 1740
Me.Top = 1995
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 4 And KeyCode = 115 Then
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
End
End If
End Sub

Private Sub Form_Load()
'On Error GoTo call_err_handle
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 548, 305, 20, 20), True

Me.Caption = "Creating User"
Me.Icon = LoadPicture(App.Path & "\Common\Icons\App Icons\FrmCreateUser.ico")
imgIcon.Picture = LoadPicture(App.Path & "\Common\Icons\App Icons\FrmCreateUser.ico")

imgPhoto.Picture = LoadPicture(App.Path & "\Common\Images\copy.jpg")
imgTheme.Picture = LoadPicture(App.Path & "\Common\Images\setting1.jpg")
accepted = False
cmbThemes.AddItem ("Gray")
cmbThemes.AddItem ("Green")
cmbThemes.AddItem ("Blue")
cmbThemes.AddItem ("Red")
cmbThemes.Text = cmbThemes.List(0)

characterlocation = App.Path & "\Chars\"
Call Agent1.Characters.Load("GENIUS", characterlocation & "GENIUS.acs")
Set cc = Agent1.Characters("GENIUS")
Call cc.MoveTo(330, 50)
Call cc.Show
Call cc.Play("Greeting")



For i = 0 To Chk1.Count - 1
Chk1(i).Enabled = False
Chk1(i).Value = 0
Next

ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide


call_err_handle:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "Information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & GetSetting(App.title, "Theme", "Name", "Gray") & "\Cursors\Move.cur")
Call ReleaseCapture
Call SendMessage(hWnd, &HA1, 2, 0&)

Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & GetSetting(App.title, "Theme", "Name", "Gray") & "\Cursors\Arrow.cur")
'*********************************
  ' hold down left mouse button and
  ' then move mouse for moving form
  '*********************************
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo call_err_handle
Call ShowButtons(Me, "000")
call_err_handle:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "Information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If

End Sub
Private Sub imgClose_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
End Sub
Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo call_err_handle

Call ShowButtons(Me, "001")
call_err_handle:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If
End Sub
Private Sub imgMin_Click()
Me.WindowState = 1
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
End Sub
Private Sub imgRestore_Click()
Me.Left = 1740
Me.Top = 1995
End Sub
Private Sub imgRestore_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "010")

q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If

End Sub

Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "100")

q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click
End If

End Sub

Private Sub txtPassword_Change()
If (txtUserName.Text <> "" And txtPassword.Text <> "" And txtVerify.Text <> "" And txtVerify.Text = txtPassword.Text) Then
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\accepted.jpg")
lblAccepted.Caption = "Accepted"
accepted = True
Else
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\notaccepted.jpg")
lblAccepted.Caption = "Not Accepted"
accepted = False
End If
End Sub

Private Sub txtUserName_Change()
lblTitle.Caption = " "
Me.Caption = "Creating User Account - " & StrConv(txtUserName.Text, vbProperCase) & "*"
lblTitle.Caption = "Creating User Account - " & StrConv(txtUserName.Text, vbProperCase) & "*"

If (txtUserName.Text <> "" And txtPassword.Text <> "" And txtVerify.Text <> "" And txtVerify.Text = txtPassword.Text) Then
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\accepted.jpg")
lblAccepted.Caption = "Accepted"
accepted = True
Else
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\notaccepted.jpg")
lblAccepted.Caption = "Not Accepted"
accepted = False
End If

End Sub
Private Sub txtUserName_GotFocus()
Call cc.Speak("Enter login name. Maximum 15 characters.")
SendKeys "{end}+{home}"
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub
Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & GetSetting(App.title, "Theme", "Name", "Gray") & "\Cursors\Arrow.cur")
txtUserName.ToolTipText = "Enter the login name here"
End Sub

Private Sub txtPassword_GotFocus()
Call cc.Speak("Enter your password. Maximum 15 characters.")
SendKeys "{end}+{home}"
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtVerify.SetFocus
End If
End Sub
Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtPassword.ToolTipText = "Enter your password here"
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & GetSetting(App.title, "Theme", "Name", "Gray") & "\Cursors\Arrow.cur")

End Sub

Private Sub txtVerify_Change()
If (txtUserName.Text <> "" And txtPassword.Text <> "" And txtVerify.Text <> "" And txtVerify.Text = txtPassword.Text) Then
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\accepted.jpg")
lblAccepted.Caption = "Accepted"
accepted = True
Else
imgAccept.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Extras\notaccepted.jpg")
lblAccepted.Caption = "Not Accepted"
accepted = False
End If
End Sub

Private Sub txtVerify_GotFocus()
Call cc.Speak("Verify your password. Maximum 15 characters.")
SendKeys "{end}+{home}"
End Sub
Private Sub txtVerify_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call click_ok
End If
End Sub
Private Sub txtVerify_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtVerify.ToolTipText = "ReEnter your password here"
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & GetSetting(App.title, "Theme", "Name", "Gray") & "\Cursors\Arrow.cur")

End Sub
Public Sub click_ok()
On Error GoTo q:
If txtUserName.Text = "" Then
'Handle_Error(title As String, err_no As String, err_desc As String, err_ico As String, err_but As Integer)
    Call cc.Speak("User name should not be empty. This is not done !")
    Call cc.Play("GestureDown")
    Call Handle_Error("Error Creating User", "Empty Login Name.", "User name should not be empty. This is not done !", "Information1.jpg", "Error.ico", 1, 0)
    ThemeIndex = cmbThemes.ListIndex
    frmMsgbox.Show vbModal
    cmbThemes.Text = cmbThemes.List(ThemeIndex)
    Call cmbThemes_Click
    txtUserName.SetFocus
Else
    If txtPassword.Text = "" Then
    Call cc.Speak("Password should not be empty. This is not done !")
    Call cc.Play("GestureDown")
     Call Handle_Error("Error Creating User", "Empty Password.", "Password should not be empty. This is not done !", "Information1.jpg", "Error.ico", 1, 0)
    ThemeIndex = cmbThemes.ListIndex
    frmMsgbox.Show vbModal
    cmbThemes.Text = cmbThemes.List(ThemeIndex)
   Call cmbThemes_Click


     txtPassword.SetFocus
    Else
        If txtVerify.Text = "" Then
        Call cc.Speak("Verify Password should not be empty. This is not done !")
        Call cc.Play("GestureDown")
        Call Handle_Error("Error Creating User", "Empty verify Password.", "Password not verified. It seems empty verify password is entered.", "Information1.jpg", "Error.ico", 1, 0)
        ThemeIndex = cmbThemes.ListIndex
        frmMsgbox.Show vbModal
        cmbThemes.Text = cmbThemes.List(ThemeIndex)
       Call cmbThemes_Click

   
        txtVerify.SetFocus
        Else
            If txtPassword.Text <> txtVerify.Text Then
            Call cc.Speak("Password does not match. Enter the same password again..")
            Call cc.Play("GestureDown")
            Call Handle_Error("Error Creating User", "Password Mismatch.", "Password does not match. Enter the same password again..", "Information1.jpg", "Error.ico", 1, 0)
            ThemeIndex = cmbThemes.ListIndex
            frmMsgbox.Show vbModal
            cmbThemes.Text = cmbThemes.List(ThemeIndex)
           Call cmbThemes_Click



            txtVerify.SetFocus
                
            Else
                Open App.Path & "\details.txt" For Output As #1
                
                If MyFile.FolderExists(App.Path & "\Tables") Then
                    MyFile.DeleteFolder (App.Path & "\Tables")
                End If
                
                MyFile.CreateFolder (App.Path & "\Tables")
               

                
                Screen.MousePointer = 99
                Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\No.cur")

                For i = 0 To 10
                Frame2.Visible = False
                Next
                Frame1.Visible = True
                Call cc.Hide
                
                Print #1, "DETAILED CONFIGURATION INFORMATION"
                Print #1, "------------------------------------------------------------------------------------------------------"
                Print #1, ""
                Dim StartTime As String
                StartTime = Format(Now, "HH:MM:SS")
                Print #1, "PROCESS STARTED AT " & StartTime & " ON " & Format(Now, "DD-MMM-YYYY")
                Print #1, ""
                
                
                lblTitle.Caption = "Updating System And Registering Components"
                Me.Caption = "Updating System And Registering Components"
                
                Call Handle_Error("Updating System.", "Updating System.", "Program will update your system. Please click Ok and wait... ", "Tip2.jpg", "Tip.ico", 1, 0)
                ThemeIndex = cmbThemes.ListIndex
                frmMsgbox.Show vbModal
                cmbThemes.Text = cmbThemes.List(ThemeIndex)
                Call cmbThemes_Click
                
                Print #1, "OPENING CONNECTION AT " & Format(Now, "HH:MM:SS")
                Print #1, ""
                                               
                'open the database
                Cn.Open "Provider=MSDAORA.1;Password=manager;User ID=system"
                On Error GoTo q1
                Chk1(0).Value = 1
                Print #1, "CONNECTION OPENED AT " & Format(Now, "HH:MM:SS")
                Print #1, ""
                
               
                errStr = "Error Creating USER"
                errDes = "Program could not create USER " & txtUserName.Text & ". May be it do exist or its an internal error. Continue anyway..."
                
                Cmd.ActiveConnection = Cn
                'CmdText = create user manohar identified by manohar
                Cmd.CommandText = "create user " & txtUserName.Text & " identified by " & txtPassword.Text
                Print #1, "CREATING USER"
                Print #1, ""
                Print #1, Cmd.CommandText
                Print #1, ""
                Cmd.Execute
                'MsgBox "user created"
                
                CmdText = "grant dba to " & txtUserName.Text
                Cmd.CommandText = CmdText
                Print #1, "GRANTING PRIVILEDGES"
                Print #1, ""
                Print #1, Cmd.CommandText
                Print #1, ""
                Cmd.Execute
                'MsgBox "dba granted"
                
                              
                Chk1(1).Value = 1
                Print #1, "USER CREATED"
                Print #1, ""
                Cn.Close
                ConctStr = "Provider=MSDAORA.1;Password=" & txtUserName.Text & ";User ID=" & txtPassword.Text
                Cn.Open ConctStr
                Cmd.ActiveConnection = Cn
                Print #1, "OPENING CONNECTION"
                Print #1, ""
                Print #1, "Provider=MSDAORA.1;Password=" & txtUserName.Text & ";User ID=" & txtPassword.Text
                Print #1, ""
                
                Print #1, "CONNECTED TO ORACLE AT " & Format(Now, "HH:MM:SS")
                Print #1, ""
                
                'MsgBox "Connection Succeeded"
                errStr = "Error Creating TABLE"
                
                'CREATE THE INSURANCE DATABASE TABLES
                Print #1, "CREATING INSURANCE DATABASE TABLES"
                Print #1, ""
                
                errDes = "Program could not create Table PERSON. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE PERSON(DRIVER_ID VARCHAR2(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(DRIVER_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE PERSON(DRIVER_ID VARCHAR2(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(DRIVER_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                
                errDes = "Program could not create Table MANUFACTURER. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE MANUFACTURER(NAME VARCHAR2(50) NOT NULL,YEAR NUMBER(4),WORTH NUMBER(12),FOUNDER VARCHAR2(25),CHAIRMAN VARCHAR2(25),ADDRESS VARCHAR2(25),CITY VARCHAR2(25),COUNTRY VARCHAR2(25),PRIMARY KEY (NAME))"
                Print #1, ""
                r = CreateTable("CREATE TABLE MANUFACTURER(NAME VARCHAR2(50) NOT NULL,YEAR NUMBER(4),WORTH NUMBER(12),FOUNDER VARCHAR2(25),CHAIRMAN VARCHAR2(25),ADDRESS VARCHAR2(25),CITY VARCHAR2(25),COUNTRY VARCHAR2(25),PRIMARY KEY (NAME))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table MANUFACTURER. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CARMAKES(NAME VARCHAR2(50) NOT NULL,VENDOR VARCHAR2(50),YEAR NUMBER(4),ADDRESS VARCHAR2(25),CITY VARCHAR2(25),COUNTRY VARCHAR2(25),PRIMARY KEY (NAME))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CARMAKES(NAME VARCHAR2(50) NOT NULL,VENDOR VARCHAR2(50),YEAR NUMBER(4),ADDRESS VARCHAR2(25),CITY VARCHAR2(25),COUNTRY VARCHAR2(25),PRIMARY KEY (NAME))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table CARMODEL. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CARMODELS(NAME VARCHAR2(25) NOT NULL,MAKE VARCHAR2(25),M_TYPE VARCHAR2(55),PRICE VARCHAR2(15),ACC0TO100 VARCHAR2(15),ACC0TO200 VARCHAR2(15),ACC0TO300 VARCHAR2(15),ACC0TO400 VARCHAR2(15),ACC0TO500 VARCHAR2(15),TOPSPEED VARCHAR2(25),AVGSPEED VARCHAR2(25),E_LAYOUT VARCHAR2(20),E_MAXPOWER VARCHAR2(25),E_MAXTORQUE VARCHAR2(25),E_SOUTPUT VARCHAR2(25),E_PTWR VARCHAR2(25),E_INSTALL VARCHAR2(25),F_AVG VARCHAR2(25),F_CITY VARCHAR2(25),F_HIGHWAY VARCHAR2(25),F_CAPACITY VARCHAR2(25),GEARBOX VARCHAR2(15),S_FRONT VARCHAR2(25),S_REAR VARCHAR2(25),STR_TYPE VARCHAR2(25),STR_POWER VARCHAR2(25),STR_TURNS VARCHAR2(25),W_SSIZE VARCHAR2(25),W_RSIZE VARCHAR2(25),W_MADEOF VARCHAR2(25),T_MODELS VARCHAR2(25),T_FRONT VARCHAR2(25),T_REAR VARCHAR2(25),B_FRONT VARCHAR2(25),B_REAR VARCHAR2(25),B_DISCS VARCHAR2(25),B_LENGTH VARCHAR2(25),B_WIDTH VARCHAR2(25),B_HEIGHT VARCHAR2(25),B_WHEELBASE VARCHAR2(25),B_FRONTTRACK VARCHAR2(25),B_REARTRACK VARCHAR2(25))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CARMODELS(NAME VARCHAR2(25) NOT NULL,MAKE VARCHAR2(25),M_TYPE VARCHAR2(55),PRICE VARCHAR2(15),ACC0TO100 VARCHAR2(15),ACC0TO200 VARCHAR2(15),ACC0TO300 VARCHAR2(15),ACC0TO400 VARCHAR2(25),ACC0TO500 VARCHAR2(25),TOPSPEED VARCHAR2(25),AVGSPEED VARCHAR2(25),E_LAYOUT VARCHAR2(20),E_MAXPOWER VARCHAR2(25),E_MAXTORQUE VARCHAR2(25),E_SOUTPUT VARCHAR2(25),E_PTWR VARCHAR2(25),E_INSTALL VARCHAR2(25),F_AVG VARCHAR2(25),F_CITY VARCHAR2(25),F_HIGHWAY VARCHAR2(25),F_CAPACITY VARCHAR2(25),GEARBOX VARCHAR2(15),S_FRONT VARCHAR2(25),S_REAR VARCHAR2(25),STR_TYPE VARCHAR2(25),STR_POWER VARCHAR2(25),STR_TURNS VARCHAR2(25),W_SSIZE VARCHAR2(25),W_RSIZE VARCHAR2(25),W_MADEOF VARCHAR2(25),T_MODELS VARCHAR2(25),T_FRONT VARCHAR2(25),T_REAR VARCHAR2(25),B_FRONT VARCHAR2(25),B_REAR VARCHAR2(25),B_DISCS VARCHAR2(25),B_LENGTH VARCHAR2(25),B_WIDTH VARCHAR2(25),B_HEIGHT VARCHAR2(25),B_WHEELBASE VARCHAR2(25),B_FRONTTRACK VARCHAR2(25),B_REARTRACK VARCHAR2(25))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                
                errDes = "Program could not create Table CARS. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CARS(REG_NO VARCHAR2(10)NOT NULL,MODEL VARCHAR2(25),YEAR NUMBER(4),CARDESC VARCHAR2(50),MAKE VARCHAR2(25), PRIMARY KEY(REG_NO))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CARS(REG_NO VARCHAR2(10)NOT NULL,MODEL VARCHAR2(25),YEAR NUMBER(4),CARDESC VARCHAR2(50),MAKE VARCHAR2(25), PRIMARY KEY(REG_NO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ACCIDENT. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ACCIDENT(REPORT_NO NUMBER(5),ACC_DATE DATE,LOCATION VARCHAR2(20), PRIMARY KEY (REPORT_NO))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ACCIDENT(REPORT_NO NUMBER(5),ACC_DATE DATE,LOCATION VARCHAR2(20), PRIMARY KEY (REPORT_NO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                                
                errDes = "Program could not create Table OWNS. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE OWNS(D_ID VARCHAR2(10) NOT NULL,REG_NUM VARCHAR2(10) NOT NULL,PRIMARY KEY (D_ID,REG_NUM),FOREIGN KEY(D_ID) REFERENCES PERSON(DRIVER_ID),FOREIGN KEY(REG_NUM) REFERENCES CARS(REG_NO))"
                Print #1, ""
                r = CreateTable("CREATE TABLE OWNS(D_ID VARCHAR2(10) NOT NULL,REG_NUM VARCHAR2(10) NOT NULL,PRIMARY KEY (REG_NUM),FOREIGN KEY(D_ID) REFERENCES PERSON(DRIVER_ID),FOREIGN KEY(REG_NUM) REFERENCES CARS(REG_NO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table PARTICIPATED. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE PARTICIPATED(DR_ID VARCHAR2(10) NOT NULL,REG_NUMB VARCHAR2(10) NOT NULL,REPORT_NUM NUMBER(5) NOT NULL, DAMAGE_AMOUNT NUMBER(10,2), PRIMARY KEY(DR_ID,REG_NUMB,REPORT_NUM) ,FOREIGN KEY(DR_ID) REFERENCES PERSON(DRIVER_ID),FOREIGN KEY(REG_NUMB) REFERENCES CARS(REG_NO),FOREIGN KEY(REPORT_NUM) REFERENCES ACCIDENT(REPORT_NO))"
                Print #1, ""
                r = CreateTable("CREATE TABLE PARTICIPATED(DR_ID VARCHAR2(10) NOT NULL,REG_NUMB VARCHAR2(10) NOT NULL,REPORT_NUM NUMBER(5) NOT NULL, DAMAGE_AMOUNT NUMBER(10,2), PRIMARY KEY(DR_ID,REG_NUMB,REPORT_NUM) ,FOREIGN KEY(DR_ID) REFERENCES PERSON(DRIVER_ID),FOREIGN KEY(REG_NUMB) REFERENCES CARS(REG_NO),FOREIGN KEY(REPORT_NUM) REFERENCES ACCIDENT(REPORT_NO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                Chk1(2).Value = 1
                Print #1, "INSURANCE DATABASE TABLES CREATED"
                Print #1, ""
                    
                    
                    
                    
                    
                    
                'CREATE THE ORDER PROCESSING DATABASE TABLES
                Print #1, "CREATING ORDER PROCESSING DATABASE TABLES"
                Print #1, ""
                errDes = "Program could not create Table CUST. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CUST(CUST_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(CUST_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CUST(CUST_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(CUST_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ITEM. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ITEM(ITEM_ID NUMBER(10),UPRICE NUMBER(10,2), PRIMARY KEY (ITEM_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ITEM(ITEM_ID NUMBER(10),UPRICE NUMBER(10,2), PRIMARY KEY (ITEM_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table WAREHOUSE. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE WAREHOUSE(WHOUSE NUMBER(10) NOT NULL,CITY VARCHAR2(20),PRIMARY KEY (WHOUSE))"
                Print #1, ""
                r = CreateTable("CREATE TABLE WAREHOUSE(WHOUSE NUMBER(10) NOT NULL,CITY VARCHAR2(20),PRIMARY KEY (WHOUSE))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ORDER_TAB. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ORDER_TAB(ORDER_NO NUMBER(10) NOT NULL,ODATE DATE,CUST_NO NUMBER(10),ORDER_AMT NUMBER(10,2),PRIMARY KEY(ORDER_NO), FOREIGN KEY(CUST_NO) REFERENCES CUST(CUST_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ORDER_TAB(ORDER_NO NUMBER(10) NOT NULL,ODATE DATE,CUST_NO NUMBER(10),ORDER_AMT NUMBER(10,2),PRIMARY KEY(ORDER_NO), FOREIGN KEY(CUST_NO) REFERENCES CUST(CUST_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ITEM. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ORDER_ITEM(I_ORDER NUMBER(10) NOT NULL,IT NUMBER(10) NOT NULL,QUT NUMBER(10),PRIMARY KEY (I_ORDER,IT),FOREIGN KEY(I_ORDER) REFERENCES ORDER_TAB(ORDER_NO),FOREIGN KEY(IT) REFERENCES ITEM(ITEM_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ORDER_ITEM(I_ORDER NUMBER(10) NOT NULL,IT NUMBER(10) NOT NULL,QUT NUMBER(10),PRIMARY KEY (I_ORDER,IT),FOREIGN KEY(I_ORDER) REFERENCES ORDER_TAB(ORDER_NO),FOREIGN KEY(IT) REFERENCES ITEM(ITEM_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                                
                errDes = "Program could not create Table SHIPMENT. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE SHIPMENT(S_ORDER NUMBER(10) NOT NULL,WHOUSENO NUMBER(10) NOT NULL,SHIP_DATE DATE,PRIMARY KEY (S_ORDER,WHOUSENO),FOREIGN KEY(S_ORDER) REFERENCES ORDER_TAB(ORDER_NO),FOREIGN KEY(WHOUSENO) REFERENCES WAREHOUSE(WHOUSE))"
                Print #1, ""
                r = CreateTable("CREATE TABLE SHIPMENT(S_ORDER NUMBER(10) NOT NULL,WHOUSENO NUMBER(10) NOT NULL,SHIP_DATE DATE,PRIMARY KEY (S_ORDER,WHOUSENO),FOREIGN KEY(S_ORDER) REFERENCES ORDER_TAB(ORDER_NO),FOREIGN KEY(WHOUSENO) REFERENCES WAREHOUSE(WHOUSE))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""

                Chk1(3).Value = 1
                Print #1, "ORDER PROCESSING DATABASE TABLES CREATED"
                Print #1, ""
                
                
                
                
                
                
                
                'CREATE STUDENT ENROLLMENT DATABASE
                Print #1, "CREATING STUDENT ENROLLMENT DATABASE TABLES"
                Print #1, ""
                
                errDes = "Program could not create Table STUDENT. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE STUDENT(REGNO VARCHAR2(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(REGNO))"
                Print #1, ""
                
                r = CreateTable("CREATE TABLE STUDENT(REGNO VARCHAR2(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(REGNO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                errDes = "Program could not create Table COURSE. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE COURSE(COURSE_ID NUMBER(2),CNAME VARCHAR2(20),DEPT VARCHAR2(20),PRIMARY KEY(COURSE_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE COURSE(COURSE_ID NUMBER(2),CNAME VARCHAR2(20),DEPT VARCHAR2(20),PRIMARY KEY(COURSE_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ENROLL. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ENROLL(REGNO VARCHAR2(10),COURSE NUMBER(2),SEM NUMBER(1),MARKS NUMBER(3),PRIMARY KEY(REGNO,COURSE,SEM),FOREIGN KEY(REGNO) REFERENCES STUDENT(REGNO),FOREIGN KEY(COURSE) REFERENCES COURSE(COURSE_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ENROLL(REGNO VARCHAR2(10),COURSE NUMBER(2),SEM NUMBER(1),MARKS NUMBER(3),PRIMARY KEY(REGNO,COURSE,SEM),FOREIGN KEY(REGNO) REFERENCES STUDENT(REGNO),FOREIGN KEY(COURSE) REFERENCES COURSE(COURSE_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table TEXT. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE TEXT(ISBN VARCHAR2(20),TITLE VARCHAR2(30),PUBLISHER VARCHAR2(30),AUTHOR VARCHAR2(50),PRIMARY KEY(ISBN))"
                Print #1, ""
                r = CreateTable("CREATE TABLE TEXT(ISBN VARCHAR2(20),TITLE VARCHAR2(30),PUBLISHER VARCHAR2(30),AUTHOR VARCHAR2(50),PRIMARY KEY(ISBN))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table BOOK_ADOPTION. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE BOOK_ADOPTION(COURSE NUMBER(2),SEM NUMBER(1),BOOK_ISBN VARCHAR2(20),PRIMARY KEY(COURSE,SEM),FOREIGN KEY(COURSE) REFERENCES COURSE(COURSE_ID),FOREIGN KEY(BOOK_ISBN) REFERENCES TEXT(ISBN))"
                Print #1, ""
                r = CreateTable("CREATE TABLE BOOK_ADOPTION(COURSE NUMBER(2),SEM NUMBER(1),BOOK_ISBN VARCHAR2(20),PRIMARY KEY(COURSE,SEM),FOREIGN KEY(COURSE) REFERENCES COURSE(COURSE_ID),FOREIGN KEY(BOOK_ISBN) REFERENCES TEXT(ISBN))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""

                Chk1(4).Value = 1
                Print #1, "STUDENT ENROLLMENT DATABASE TABLES CREATED"
                Print #1, ""
                
                
                
                
                
                
                
                
                'CREATE BOOK DELER DATABASE
                Print #1, "CREATING BOOK DELER DATABASE TABLES"
                Print #1, ""
                
                errDes = "Program could not create Table AUTHOR. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE AUTHOR(AUTHOR_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),COUNTRY VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(AUTHOR_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE AUTHOR(AUTHOR_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),COUNTRY VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(AUTHOR_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table PUBLISHER. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE PUBLISHER(PUB_ID NUMBER(10),NAME VARCHAR2(25),CITY VARCHAR2(20),COUNTRY VARCHAR2(20),PRIMARY KEY(PUB_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE PUBLISHER(PUB_ID NUMBER(10),NAME VARCHAR2(25),CITY VARCHAR2(20),COUNTRY VARCHAR2(20),PRIMARY KEY(PUB_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table CATAGORY. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CATEGORY(CATEGORY NUMBER(4),DESCRIPTION VARCHAR2(50),PRIMARY KEY(CATEGORY))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CATEGORY(CATEGORY NUMBER(4),DESCRIPTION VARCHAR2(50),PRIMARY KEY(CATEGORY))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table CATALOG. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CATALOG(BOOK_ID NUMBER(4),TITLE VARCHAR2(25),AUTHOR_ID NUMBER(10),PUBLISH_ID NUMBER(10),CATEGORY_ID NUMBER(4),YEAR NUMBER(4),PRICE NUMBER(10,2),PRIMARY KEY(BOOK_ID),FOREIGN KEY(AUTHOR_ID) REFERENCES AUTHOR(AUTHOR_ID),FOREIGN KEY(PUBLISH_ID) REFERENCES PUBLISHER(PUB_ID),FOREIGN KEY(CATEGORY_ID) REFERENCES CATEGORY(CATEGORY))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CATALOG(BOOK_ID NUMBER(4),TITLE VARCHAR2(25),AUTHOR_ID NUMBER(10),PUBLISH_ID NUMBER(10),CATEGORY_ID NUMBER(4),YEAR NUMBER(4),PRICE NUMBER(10,2),PRIMARY KEY(BOOK_ID),FOREIGN KEY(AUTHOR_ID) REFERENCES AUTHOR(AUTHOR_ID),FOREIGN KEY(PUBLISH_ID) REFERENCES PUBLISHER(PUB_ID),FOREIGN KEY(CATEGORY_ID) REFERENCES CATEGORY(CATEGORY))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ORDER_DETAILS. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ORDER_DETAILS(ORDER_NO NUMBER(4),BOOK_ID NUMBER(10),QTY NUMBER(4),PRIMARY KEY(ORDER_NO,BOOK_ID),FOREIGN KEY(BOOK_ID) REFERENCES CATALOG(BOOK_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ORDER_DETAILS(ORDER_NO NUMBER(4),BOOK_ID NUMBER(10),QTY NUMBER(4),PRIMARY KEY(ORDER_NO,BOOK_ID),FOREIGN KEY(BOOK_ID) REFERENCES CATALOG(BOOK_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                Chk1(5).Value = 1
                Print #1, "BOOK DELER DATABASE TABLES CREATED"
                Print #1, ""
                
                
                
                'CREATE BANKING INTERPRISE DATABASE
                Print #1, "CREATING BANKING INTERPRISE DATABASE TABLES"
                Print #1, ""
                
                errDes = "Program could not create Table BRANCH. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE BRANCH(NAME VARCHAR2(20),CITY VARCHAR2(20),ASSETS NUMBER(12,2),PRIMARY KEY(NAME))"
                Print #1, ""
                r = CreateTable("CREATE TABLE BRANCH(NAME VARCHAR2(20),CITY VARCHAR2(20),ASSETS NUMBER(12,2),PRIMARY KEY(NAME))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table ACCOUNT. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE ACCOUNT(ACCNO NUMBER(10),BRANCH_NAME VARCHAR2(20),BALANCE NUMBER(12,2),PRIMARY KEY(ACCNO),FOREIGN KEY(BRANCH_NAME) REFERENCES BRANCH(NAME))"
                Print #1, ""
                r = CreateTable("CREATE TABLE ACCOUNT(ACCNO NUMBER(10),BRANCH_NAME VARCHAR2(20),BALANCE NUMBER(12,2),PRIMARY KEY(ACCNO),FOREIGN KEY(BRANCH_NAME) REFERENCES BRANCH(NAME))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table CUSTOMER. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE CUSTOMER(CUST_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(CUST_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE CUSTOMER(CUST_ID NUMBER(10) NOT NULL,FNAME VARCHAR2(20) NOT NULL,MNAME VARCHAR2(15),LNAME VARCHAR2(15),SEX VARCHAR2(6),DOB DATE,ADDRESS VARCHAR2(30),CITY VARCHAR2(30),DISTT VARCHAR2(30),STATE VARCHAR2(30),PIN VARCHAR2(6),PHONE VARCHAR2(15),PHOTO VARCHAR2(100), PRIMARY KEY(CUST_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table DEPOSITOR. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE DEPOSITOR(ACCNO NUMBER(10),CUST_ID NUMBER(10),PRIMARY KEY(ACCNO,CUST_ID),FOREIGN KEY(ACCNO) REFERENCES ACCOUNT(ACCNO),FOREIGN KEY(CUST_ID) REFERENCES CUSTOMER(CUST_ID))"
                Print #1, ""
                r = CreateTable("CREATE TABLE DEPOSITOR(ACCNO NUMBER(10),CUST_ID NUMBER(10),PRIMARY KEY(ACCNO,CUST_ID),FOREIGN KEY(ACCNO) REFERENCES ACCOUNT(ACCNO),FOREIGN KEY(CUST_ID) REFERENCES CUSTOMER(CUST_ID))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table LOAN. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE LOAN(LOAN_NO NUMBER(10),B_NAME VARCHAR2(20),AMT NUMBER(12,2),PRIMARY KEY(LOAN_NO),FOREIGN KEY(B_NAME) REFERENCES BRANCH(NAME))"
                Print #1, ""
                r = CreateTable("CREATE TABLE LOAN(LOAN_NO NUMBER(10),B_NAME VARCHAR2(20),AMT NUMBER(12,2),PRIMARY KEY(LOAN_NO),FOREIGN KEY(B_NAME) REFERENCES BRANCH(NAME))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                errDes = "Program could not create Table BORROWER. May be it do exist or its an internal error. Continue anyway..."
                Print #1, "CREATE TABLE BORROWER(C_NAME NUMBER(10),LOAN_NO NUMBER(10),PRIMARY KEY(C_NAME,LOAN_NO),FOREIGN KEY(C_NAME) REFERENCES CUSTOMER(CUST_ID),FOREIGN KEY(LOAN_NO) REFERENCES LOAN(LOAN_NO))"
                Print #1, ""
                r = CreateTable("CREATE TABLE BORROWER(C_NAME NUMBER(10),LOAN_NO NUMBER(10),PRIMARY KEY(C_NAME,LOAN_NO),FOREIGN KEY(C_NAME) REFERENCES CUSTOMER(CUST_ID),FOREIGN KEY(LOAN_NO) REFERENCES LOAN(LOAN_NO))")
                Print #1, "Table - ' " & Mid(r, 2, Len(r)) & " ' CREATED"
                Print #1, ""
                
                
                Chk1(6).Value = 1
                Print #1, "BANKING INTERPRISE DATABASE TABLES CREATED"
                Print #1, ""
                
                'this goes to windows registry
                Print #1, "UPDATING WINDOWS REGISTRY"
                Print #1, ""
                Print #1, "------------------------------------------------------------------------------------------------------"
                Print #1, ""
                
               
                SaveSetting App.title, "Users", "Name", txtUserName.Text
                SaveSetting App.title, "Users", "PassWord", txtPassword.Text
                Print #1, "USER NAME " & txtUserName.Text
                Print #1, ""
                Print #1, "PASSWORD  " & txtPassword.Text
                Print #1, ""
                
                
                
                Chk1(7).Value = 1
                On Error GoTo q
                SaveSetting App.title, "Theme", "Name", cmbThemes.Text
                Call Save_Theme
                Print #1, "THEME NAME " & cmbThemes.Text
                Print #1, ""
                               
                SaveSetting App.title, "Settings", "AutoLogin", "No"
                Print #1, "AUTO LOAD  " & "No"
                Print #1, ""
                
                Chk1(8).Value = 1
                Print #1, "WINDOWS REGISTRY UPDATED"
                Print #1, ""
                Print #1, "------------------------------------------------------------------------------------------------------"
                Print #1, ""
                'delete any temporary files
                
                Chk1(9).Value = 1
                
                lblTitle.Caption = "System Updated"
                Me.Caption = "System Updated"
                Call Handle_Error("System Updated.", "Updation Complete.", "Your system updated correctly. You can continue now.", "Tip2.jpg", "tip.ico", 1, 0)
                frmMsgbox.Show vbModal
                cmbThemes.Text = cmbThemes.List(0)
                Call cmbThemes_Click


                cmdShowDetail.Enabled = True
                cmdOkContinue.Enabled = True
                cmdOkContinue.SetFocus
                
                'close the connection
                Cn.Close
          
            End If
        End If
    End If
End If
                Print #1, "PROCESS STARTED AT   - " & StartTime & "    ON  " & Format(Now, "DD-MMM-YYYY")
                Print #1, ""
                Print #1, "PROCESS COMPLETED AT - " & Format(Now, "HH:MM:SS") & "    ON  " & Format(Now, "DD-MMM-YYYY")
                Print #1, ""
                
Print #1, "********************************************************************************"
Print #1, ""
                Print #1, "APPLICATION NAME     : " & App.title
                Print #1, ""
                Print #1, "APPLICATION PATH     : " & App.Path
                Print #1, ""
Print #1, "********************************************************************************"
Close #1

'

q:
If Err.Number <> 0 Then
Call Handle_Error("Runtime Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
If Err.Number = 52 Then
    Open App.Path & "\Details.txt" For Output As #1
End If
Print #1, Err.Number & Err.Description
Print #1, ""
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
cmbThemes.Text = cmbThemes.List(ThemeIndex)
Call cmbThemes_Click

cmdShowDetail.Enabled = True
cmdExitNow.Enabled = True
End If

q1:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr(errStr), CStr(errDes), "Information1.jpg", "information.ico", 2, 0)
Print #1, Err.Number & Err.Description
Print #1, ""
ThemeIndex = cmbThemes.ListIndex
frmMsgbox.Show vbModal
Call Apply_Theme(Me, 1)
If MsgBOx_R_Value Then
Resume Next
Exit Sub
Else
cmdShowDetail.Enabled = True
cmdExitNow.Enabled = True
Call Apply_Theme(Me, 1)
cmdExitNow.SetFocus
End If
End If


End Sub
