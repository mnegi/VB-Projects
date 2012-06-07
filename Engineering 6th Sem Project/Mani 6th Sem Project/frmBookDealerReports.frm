VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmBookDealerReports 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmBookDealerReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   660
      TabIndex        =   17
      Top             =   7020
      Width           =   10755
      Begin VB.ComboBox cmbNos 
         Height          =   315
         Left            =   8295
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "1"
         Top             =   180
         Width           =   2400
      End
      Begin ManoharButton.MyButton MyButton5 
         Height          =   375
         Left            =   9225
         TabIndex        =   19
         Tag             =   "1"
         Top             =   795
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Print"
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
         MICON           =   "frmBookDealerReports.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "List the item that has appeared in more than "" ? "" orders"
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
         Height          =   870
         Left            =   195
         TabIndex        =   20
         Tag             =   "1"
         Top             =   135
         Width           =   7710
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   660
      TabIndex        =   13
      Top             =   5403
      Width           =   10755
      Begin VB.ComboBox cmbDays 
         Height          =   315
         Left            =   8295
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "1"
         Top             =   180
         Width           =   2400
      End
      Begin ManoharButton.MyButton MyButton2 
         Height          =   375
         Left            =   9225
         TabIndex        =   15
         Tag             =   "1"
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Print"
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
         MICON           =   "frmBookDealerReports.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "List the details of Orders that were not shipped within "" ? "" days. ( Select the days from the list .)"
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
         Height          =   870
         Left            =   195
         TabIndex        =   16
         Tag             =   "1"
         Top             =   135
         Width           =   7710
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   660
      TabIndex        =   11
      Top             =   2456
      Width           =   10755
      Begin ManoharButton.MyButton MyButton4 
         Height          =   360
         Left            =   9240
         TabIndex        =   3
         Tag             =   "1"
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Print"
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
         MICON           =   "frmBookDealerReports.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Produce a listing  :  CustomerName,No.OfOrders,AverageOrderAmount"
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
         Height          =   300
         Left            =   195
         TabIndex        =   12
         Tag             =   "1"
         Top             =   360
         Width           =   8580
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   660
      TabIndex        =   9
      Top             =   1260
      Width           =   10755
      Begin VB.ComboBox cmbTables 
         Height          =   315
         Left            =   4815
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "1"
         Top             =   285
         Width           =   3930
      End
      Begin ManoharButton.MyButton cmbPrintT 
         Height          =   360
         Left            =   9210
         TabIndex        =   2
         Tag             =   "1"
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Pri&nt"
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
         MICON           =   "frmBookDealerReports.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print the contents of the follwing table."
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
         Left            =   165
         TabIndex        =   10
         Tag             =   "1"
         Top             =   315
         Width           =   3780
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   660
      TabIndex        =   7
      Top             =   3787
      Width           =   10755
      Begin VB.ComboBox cmbCity 
         Height          =   315
         Left            =   8265
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "1"
         Top             =   180
         Width           =   2400
      End
      Begin ManoharButton.MyButton MyButton1 
         Height          =   375
         Left            =   9210
         TabIndex        =   5
         Tag             =   "1"
         Top             =   765
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Print"
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
         MICON           =   "frmBookDealerReports.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBookDealerReports.frx":0098
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
         Height          =   870
         Left            =   195
         TabIndex        =   8
         Tag             =   "1"
         Top             =   135
         Width           =   7710
      End
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   6
      Top             =   465
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   291
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Dealer DataBase - Report Generator"
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
      Width           =   4215
   End
End
Attribute VB_Name = "frmBookDealerReports"
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

'Private Sub btnQ4_Click()
'On Error Resume Next
'    DataEnvironment1.rscmbPersonQ4.Close
'    DataEnvironment1.cmbPersonQ4 Format(cmbAdate.Text, "dd-mmm-yy")
'    dsrPersonQuery4.Caption = "Report Generator - Query 4"
'    dsrPersonQuery4.WindowState = 2
'    dsrPersonQuery4.Show
'End Sub

Private Sub cmbPrintT_Click()
On Error Resume Next
Select Case cmbTables.Text
Case "Author"
    DataEnvironment1.rscmdAuthor.Close
    DataEnvironment1.cmdAuthor
    dsrAuthor.Caption = "Report Generator - Author Table"
    dsrAuthor.WindowState = 2
    dsrAuthor.Show
Case "Publisher"
    DataEnvironment1.rscmdPublisher.Close
    DataEnvironment1.cmdPublisher
    dsrPublisher.Caption = "Report Generator - Publisher Table"
    dsrPublisher.WindowState = 2
    dsrPublisher.Show
Case "Catalog"
    DataEnvironment1.rscmdCatalog.Close
    DataEnvironment1.cmdCatalog
    dsrCatalog.Caption = "Report Generator - Catalog Table"
    dsrCatalog.WindowState = 2
    dsrCatalog.Show
Case "Catagory"
    DataEnvironment1.rscmdCatagory.Close
    DataEnvironment1.cmdCatagory
    dsrCatagory.Caption = "Report Generator - Catagory Table"
    dsrCatagory.WindowState = 2
    dsrCatagory.Show
Case "Order Details"
    DataEnvironment1.rscmdOrderDetails.Close
    DataEnvironment1.cmdOrderDetails
    dsrOrderD.Caption = "Report Generator - Order Details Table"
    dsrOrderD.WindowState = 2
    dsrOrderD.Show

End Select
End Sub

Private Sub Form_Activate()
'On Error GoTo q
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
'On Error GoTo q
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
Me.Caption = "Book Dealer DataBase - Report Generator"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")


Call loadTables

Dim rsD As New ADODB.Recordset
rsD.Open "select distinct city from warehouse", Cn, adOpenStatic
If rsD.EOF = True And rsD.BOF = True Then
'
Else
rsD.MoveFirst
cmbCity.CLEAR
While Not rsD.EOF
    cmbCity.AddItem rsD.Fields(0)
    rsD.MoveNext
Wend
cmbCity.Text = cmbCity.List(0)
End If
rsD.Close

cmbDays.CLEAR
cmbNos.CLEAR
For i = 1 To 366
    cmbDays.AddItem i
    cmbNos.AddItem i
Next
cmbDays.Text = cmbDays.List(0)
cmbNos.Text = cmbNos.List(0)


q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
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
'On Error GoTo q
Call ShowButtons(Me, "010")
q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
Unload Me
frmBookMain.Show
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
Call ShowButtons(Me, "001")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub imgMin_Click()
'On Error GoTo q
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
'On Error GoTo q
Call ShowButtons(Me, "100")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub
Private Sub MyButton1_Click()
On Error Resume Next
    DataEnvironment1.rscmd2_4.Close
    DataEnvironment1.cmd2_4 cmbCity.Text
    dsr2_4.Caption = "Report Generator - Orders that were shipped from " & cmbCity.Text
    dsr2_4.WindowState = 2
    dsr2_4.Show
    
End Sub

Private Sub MyButton2_Click()
On Error Resume Next
    DataEnvironment1.rscmd2E1.Close
    DataEnvironment1.cmd2E1 CInt(cmbDays.Text)
    dsr2E1.Caption = "Report Generator - Orders that were not shipped with in " & cmbDays.Text & " days."
    dsr2E1.WindowState = 2
'    dsr2E1.Label3.Caption = "Report Generator - Orders that were not shipped with in " & cmbDays.Text & " days."
    dsr2E1.Show

End Sub

Private Sub MyButton3_Click()
On Error Resume Next
    DataEnvironment1.rscmd1E_1.Close
    DataEnvironment1.cmd1E_1
    dsrInsEQ1.Caption = "Report Generator - Owners who were never involved in any accident."
    dsrInsEQ1.WindowState = 2
    dsrInsEQ1.Show
End Sub

Private Sub MyButton4_Click()
On Error Resume Next
    DataEnvironment1.rscmd2_3.Close
    DataEnvironment1.cmd2_3
    dsr2_3.Caption = "Report Generator - Listing Of CustomerId,NumberOfOrders and Average Order Amount"
    dsr2_3.WindowState = 2
    dsr2_3.Show
End Sub
Sub loadTables()
cmbTables.CLEAR
cmbTables.AddItem "Author"
cmbTables.AddItem "Publisher"
cmbTables.AddItem "Catalog"
cmbTables.AddItem "Catagory"
cmbTables.AddItem "Order Details"
cmbTables.Text = cmbTables.List(0)
End Sub

Private Sub MyButton5_Click()
On Error Resume Next
    DataEnvironment1.rscmd2E2.Close
    DataEnvironment1.cmd2E2 CInt(cmbNos.Text)
    dsr2E2.Caption = "Report Generator - Listing of items  that have been appeared in more than " & cmbNos.Text & " orders."
    dsr2E2.WindowState = 2
    dsr2E2.Show
End Sub
