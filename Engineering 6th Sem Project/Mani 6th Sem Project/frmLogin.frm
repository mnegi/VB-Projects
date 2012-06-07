VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Security Validation"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   ForeColor       =   &H00404040&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   StartUpPosition =   2  'CenterScreen
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   300
      Left            =   150
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
   End
   Begin VB.CheckBox ChkAutoLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AutoLoad"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   225
      MaskColor       =   &H00404000&
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Clicking this will not promt for login onwards."
      Top             =   2175
      Width           =   1740
   End
   Begin XPTEXTBOX.text txtUserName 
      Height          =   390
      Left            =   2280
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Enter Login Name"
      Top             =   990
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   688
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      MaxLength       =   15
      LineColor       =   11643476
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
   Begin XPTEXTBOX.text txtPassword 
      Height          =   390
      Left            =   2280
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Enter Password"
      Top             =   1620
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   688
      FontName        =   "Arial"
      FontSize        =   9.75
      MaxLength       =   15
      LineColor       =   11643476
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   2175
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Login"
      Top             =   2265
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "L&ogin"
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
      MICON           =   "frmLogin.frx":000C
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
      Left            =   3675
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Exit Application"
      Top             =   2250
      Width           =   1380
      _ExtentX        =   2434
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
      MICON           =   "frmLogin.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgRestore 
      Height          =   270
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   4785
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   285
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   4455
      Stretch         =   -1  'True
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Left            =   105
      Stretch         =   -1  'True
      ToolTipText     =   "Application Icon"
      Top             =   45
      Width           =   375
   End
   Begin VB.Label lblUname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
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
      Left            =   360
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1020
      Width           =   1080
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
      Left            =   360
      TabIndex        =   6
      Tag             =   "1"
      Top             =   1605
      Width           =   945
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Varification"
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
      Left            =   1485
      TabIndex        =   1
      Tag             =   "1"
      Top             =   105
      Width           =   1980
   End
End
Attribute VB_Name = "frmLogin"
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

Private Sub ChkAutoLoad_GotFocus()
ChkAutoLoad.ToolTipText = "Don't Show Login Form Onwards"
End Sub

Private Sub ChkAutoLoad_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ChkAutoLoad.ToolTipText = "Don't Show Login Form Onwards"
End Sub

Private Sub cmdCreate_Click()
Call click_ok
End Sub
Private Sub cmExit_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
End
End Sub
Private Sub Form_Activate()
On Error GoTo q:

ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide
txtUserName.SetFocus

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), "File Not Found", CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MousePointer = 15
'Call ReleaseCapture
'Call SendMessage(hWnd, &HA1, 2, 0&)
'MousePointer = 1
'  '*********************************
'  ' hold down left mouse button and
'  ' then move mouse for moving form
'  '*********************************
'End Sub
'
Private Sub Form_Load()
On Error GoTo q1:
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

ConctStr = "Provider=MSDAORA.1;Password=" & GetSetting(App.title, "Users", "Name") & ";User ID=" & GetSetting(App.title, "Users", "Password")
'ConctStr = "Provider=MSDAORA.1;Password=scott;User ID=tiger"
Cn.Open ConctStr
Cn.Open

DataEnvironment1.Connection1.Open ConctStr
On Error GoTo q
If GetSetting(App.title, "Settings", "AutoLogin") = "No" Then
Call Get_Theme
Call Apply_Theme(Me, 0)


Me.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\small1.jpg")
Me.Icon = LoadPicture(App.Path & "\Common\Icons\App Icons\FrmLogin.ico")
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Security Validation"
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 348, 220, 20, 20), True

imgIcon.Picture = LoadPicture(App.Path & "\Common\Icons\App Icons\FrmLogin.ico")
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\C_up.jpg")
imgMin.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\M_up.jpg")

txtPassword.FontName = "Webdings"
txtPassword.FontSize = 10
txtPassword.PasswordChar = "a"
Else
Unload Me
frmMain.Show
End If

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), "File Not Found", CStr(Err.Description), "Information1.jpg", "Information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
q1:
If Err.Number <> 0 Then
Call ManiExtras1.DesktopIconsShow
Call ManiExtras1.TaskBarShow
Call Handle_Error("Error : " & Err.Number, "Oracle Not Started", "Please Start Oracle.", "Information1.jpg", "Information.ico", 1, 0)
frmMsgbox.Show vbModal
End
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q:
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\C_up.jpg")
imgMin.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\M_up.jpg")

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), "File Not Found.", CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
End
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q:
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\C_down.jpg")

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), "File Not Found", CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgMin_Click()
Me.WindowState = 1
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
End Sub

Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q:
imgMin.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\M_down.jpg")

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), "File Not Found", CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub txtPassword_GotFocus()
SendKeys "{end}+{home}"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call click_ok
End If
End Sub

Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtPassword.ToolTipText = "Enter your password here"
End Sub

Private Sub txtUserName_GotFocus()
SendKeys "{end}+{home}"
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub

Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtUserName.ToolTipText = "Enter the login name here"
End Sub

Public Sub click_ok()

If txtUserName.Text = GetSetting(App.title, "Users", "Name") Then
    If txtPassword.Text = GetSetting(App.title, "Users", "Password") Then
        If ChkAutoLoad.Value = 1 Then
            SaveSetting App.title, "Settings", "AutoLogin", "Yes"
        Else
            SaveSetting App.title, "Settings", "AutoLogin", "No"
        End If
        Unload Me
        frmMain.Show
    Else
    
    Call Handle_Error("Login Error", "Incorrect Password", "Password entered is not coreect. Please try entering the correct password.", "Information1.jpg", "Error.ico", 1, 0)
    frmMsgbox.Show vbModal
    txtPassword.SetFocus
    End If
Else
    Call Handle_Error("Login Error", "Incorrect Login Name", "Login Name entered is not coreect. Please try entering the correct login name.", "Information1.jpg", "Error.ico", 1, 0)
    frmMsgbox.Show vbModal
    txtUserName.SetFocus
    End If

End Sub

