VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmMsgbox 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Security Validation"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   ForeColor       =   &H00404040&
   Icon            =   "frmMsgbox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   StartUpPosition =   2  'CenterScreen
   Begin XPTEXTBOX.text txtMsgbox 
      Height          =   375
      Left            =   1755
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1515
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      LineColor       =   11643476
      Text            =   ""
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   300
      Left            =   150
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
   End
   Begin ManoharButton.MyButton cmdok 
      Height          =   405
      Left            =   1110
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Ok Continue AnyWay"
      Top             =   2415
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Ok"
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
      MICON           =   "frmMsgbox.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton cmdCancel 
      Height          =   405
      Left            =   2850
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Cancel This"
      Top             =   2415
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Cancel"
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
      MICON           =   "frmMsgbox.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton cmdOkOnly 
      Height          =   405
      Left            =   2010
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "That's Allright"
      Top             =   2415
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Ok"
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
      MICON           =   "frmMsgbox.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDesc 
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
      Height          =   1365
      Left            =   1680
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Error Description"
      Top             =   1050
      Width           =   3525
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   1035
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   1395
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   405
      Stretch         =   -1  'True
      ToolTipText     =   "Minimise"
      Top             =   1395
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgRestore 
      Height          =   270
      Left            =   735
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   1395
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgAppIcon 
      Height          =   330
      Left            =   150
      Stretch         =   -1  'True
      ToolTipText     =   "Application Icon"
      Top             =   45
      Width           =   315
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Error Name"
      Top             =   705
      Width           =   5100
   End
   Begin VB.Label lblTitle 
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
      Left            =   780
      TabIndex        =   0
      Tag             =   "1"
      Top             =   45
      Width           =   1980
   End
   Begin VB.Image imgIcon 
      Height          =   1485
      Left            =   120
      Stretch         =   -1  'True
      Top             =   885
      Width           =   1500
   End
End
Attribute VB_Name = "frmMsgbox"
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

Private Sub cmdCancel_Click()
MsgBOx_R_Value = False
Unload Me
End Sub
Private Sub cmdok_Click()
MsgBOx_R_Value = True
Unload Me
End Sub

Private Sub cmdOkOnly_Click()
MsgBOx_R_Value = True
Unload Me
End Sub

Private Sub Form_Activate()
Call Apply_Theme(Me, 0)
lblError.ForeColor = vbBlack
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 15
Call ReleaseCapture
Call SendMessage(hWnd, &HA1, 2, 0&)
MousePointer = 1
  '*********************************
  ' hold down left mouse button and
  ' then move mouse for moving form
  '*********************************
End Sub

Private Sub Form_Load()
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 348, 220, 20, 20), True
End Sub
