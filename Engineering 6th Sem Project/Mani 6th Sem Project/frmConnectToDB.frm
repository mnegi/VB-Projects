VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmExportData 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Craeting User Account - *"
   ClientHeight    =   4545
   ClientLeft      =   1740
   ClientTop       =   1995
   ClientWidth     =   8205
   Icon            =   "frmConnectToDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4545
   ScaleWidth      =   8205
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   270
      Left            =   1395
      TabIndex        =   6
      Top             =   465
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   476
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   270
      TabIndex        =   7
      Top             =   885
      Width           =   7740
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2340
         TabIndex        =   11
         Top             =   255
         Width           =   2535
         Begin VB.Label lblThemeSelect 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connect to"
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
            Height          =   285
            Left            =   660
            TabIndex        =   12
            Tag             =   "1"
            Top             =   180
            Width           =   1080
         End
      End
      Begin XPTEXTBOX.text txtUserName 
         Height          =   390
         Left            =   4755
         TabIndex        =   0
         Tag             =   "1"
         ToolTipText     =   "Enter the User Id"
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   ""
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
      Begin ManoharButton.MyButton cmdConnect 
         Height          =   405
         Left            =   4770
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Create User"
         Top             =   1965
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Connect"
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
         MICON           =   "frmConnectToDB.frx":000C
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
         Left            =   6375
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Exit Application"
         Top             =   1995
         Width           =   1305
         _ExtentX        =   2302
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
         MICON           =   "frmConnectToDB.frx":0028
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
         Left            =   4755
         TabIndex        =   1
         Tag             =   "2"
         ToolTipText     =   "Enter the Password"
         Top             =   690
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "Webdings"
         FontSize        =   9.75
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   ""
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
         Left            =   4755
         TabIndex        =   2
         Tag             =   "2"
         ToolTipText     =   "Enter the host string"
         Top             =   1350
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   688
         FontName        =   "Webdings"
         FontSize        =   9.75
         MaxLength       =   15
         LineColor       =   11643476
         Text            =   ""
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1830
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   2535
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Oracle"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   135
            TabIndex        =   16
            Top             =   1290
            Width           =   2265
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SQL Server"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            TabIndex        =   15
            Top             =   750
            Width           =   2265
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Microsoft Access"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   120
            TabIndex        =   14
            Top             =   210
            Width           =   2265
         End
      End
      Begin XPTEXTBOX.text txtConnectionStr 
         Height          =   450
         Left            =   30
         TabIndex        =   17
         Tag             =   "1"
         ToolTipText     =   "Enter the Password"
         Top             =   2730
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   794
         FontName        =   "Webdings"
         FontSize        =   9.75
         LineColor       =   11643476
         Text            =   ""
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection String"
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
         Height          =   285
         Left            =   45
         TabIndex        =   18
         Tag             =   "1"
         Top             =   2430
         Width           =   1785
      End
      Begin VB.Label lblHostDtring 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Host String"
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
         Height          =   285
         Left            =   2895
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1345
         Width           =   1140
      End
      Begin VB.Label lblUname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Id"
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
         Height          =   285
         Left            =   2895
         TabIndex        =   9
         Tag             =   "1"
         Top             =   45
         Width           =   750
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   285
         Left            =   2865
         TabIndex        =   8
         Tag             =   "1"
         Top             =   695
         Width           =   945
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
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manohar Export Data Wizard 1.0"
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
      TabIndex        =   5
      Tag             =   "1"
      Top             =   60
      Width           =   3225
   End
End
Attribute VB_Name = "frmExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn1 As New ADODB.Connection

Private Sub cmExit_Click()

End Sub

Private Sub Form_Activate()
On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 1)
ManiExtras1.MinimizeAll
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo q
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Move.cur")

'MousePointer = 15
Call ReleaseCapture
Call SendMessage(hWnd, &HA1, 2, 0&)
'MousePointer = 1
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")

  '*********************************
  ' hold down left mouse button and
  ' then move mouse for moving form
  '*********************************
  
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


  
End Sub

Private Sub Form_Load()
On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Connection Setup 2 OLEDB Providers"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 548, 305, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
Call Load_PasswordChar(txtPassword)
Option3.Value = True

txtConnectionStr.Text = "Provider=MSDAORA.1;User ID=?;Password=?"
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "000")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub
Private Sub imgRestore_Click()
ManiExtras1.DesktopIconsHide
ManiExtras1.TaskBarHide
Me.Left = 0
Me.Top = 0
End Sub
Private Sub imgRestore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "010")
q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "001")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub imgMin_Click()
On Error GoTo q
Me.WindowState = 1
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "100")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub
