VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmOrderMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Order Processing  DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmOrderMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   8
      Top             =   465
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   291
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   0
      Left            =   570
      TabIndex        =   2
      Tag             =   "1"
      Top             =   2145
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Customer Manager"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   345
      Top             =   1440
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   1
      Left            =   570
      TabIndex        =   3
      Tag             =   "1"
      Top             =   2845
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Order"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   2
      Left            =   570
      TabIndex        =   4
      Tag             =   "1"
      Top             =   3545
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Item"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   3
      Left            =   570
      TabIndex        =   5
      Tag             =   "1"
      Top             =   5645
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Shipment"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   4
      Left            =   570
      TabIndex        =   6
      Tag             =   "1"
      Top             =   6345
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Re&port Generator"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   7
      Left            =   570
      TabIndex        =   7
      Tag             =   "1"
      Top             =   7050
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Back To Main"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   10
      Left            =   570
      TabIndex        =   9
      Tag             =   "1"
      Top             =   4245
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "O&rder Items"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   11
      Left            =   570
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4945
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&WareHouse"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMain.frx":00D0
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
      Left            =   11265
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   75
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   60
      Width           =   330
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   10950
      Stretch         =   -1  'True
      Top             =   75
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   11595
      Stretch         =   -1  'True
      Top             =   75
      Width           =   285
   End
   Begin VB.Label lbltime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   135
      Left            =   11775
      TabIndex        =   1
      Top             =   405
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Processing DataBase Launcher"
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
      Left            =   900
      TabIndex        =   0
      Tag             =   "1"
      Top             =   75
      Width           =   3735
   End
End
Attribute VB_Name = "frmOrderMain"
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

Private Sub Btns_Click(Index As Integer)
On Error GoTo q
Select Case Index
Case 0
Me.Hide
frmCustomer.Show
Case 1
Me.Hide
frmOrder.Show
Case 2
Me.Hide
frmItem.Show
Case 3
Me.Hide
frmShipment.Show
Case 4
Me.Hide
frmOrderReports.Show

Case 10
Me.Hide
frmOrderItem.Show
Case 11
Me.Hide
frmWarehouse.Show

Case 7
Me.Hide
frmMain.Show
End Select
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Private Sub Form_Activate()
On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 3)
ManiExtras1.MinimizeAll
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Order Processing DataBase Launcher"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
'imgWallpaper.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Wallpapers\FrmMainWallper.jpg")
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Private Sub imgRestore_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "010")
q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
End
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "100")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format(Now, "hh:mm:ss", vbUseSystemDayOfWeek)
End Sub
