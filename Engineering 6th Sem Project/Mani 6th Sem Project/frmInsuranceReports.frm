VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmInsuranceRoports 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmInsuranceReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   555
      TabIndex        =   20
      Top             =   2257
      Width           =   10785
      Begin VB.ComboBox cmbA2 
         Height          =   315
         Left            =   9030
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "1"
         Top             =   195
         Width           =   1455
      End
      Begin ManoharButton.MyButton MyButton4 
         Height          =   375
         Left            =   9030
         TabIndex        =   4
         Tag             =   "1"
         Top             =   900
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
         MICON           =   "frmInsuranceReports.frx":000C
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
         Caption         =   "Find the total number of people who owned cars and were involved in accidents with their own car in year"
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
         Height          =   840
         Left            =   195
         TabIndex        =   21
         Tag             =   "1"
         Top             =   150
         Width           =   8580
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   540
      TabIndex        =   18
      Top             =   7275
      Width           =   10755
      Begin ManoharButton.MyButton MyButton3 
         Height          =   360
         Left            =   9030
         TabIndex        =   9
         Tag             =   "1"
         Top             =   285
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "Pr&int"
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
         MICON           =   "frmInsuranceReports.frx":0028
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owners who were never involved in any accidents."
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
         TabIndex        =   19
         Tag             =   "1"
         Top             =   315
         Width           =   4950
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   555
      TabIndex        =   15
      Top             =   1170
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
         Left            =   9030
         TabIndex        =   2
         Tag             =   "1"
         Top             =   285
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
         MICON           =   "frmInsuranceReports.frx":0044
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
         TabIndex        =   16
         Tag             =   "1"
         Top             =   315
         Width           =   3780
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   540
      TabIndex        =   13
      Top             =   3929
      Width           =   10785
      Begin VB.ComboBox cmbA 
         Height          =   315
         Left            =   9030
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1"
         Top             =   195
         Width           =   1455
      End
      Begin ManoharButton.MyButton MyButton1 
         Height          =   375
         Left            =   9030
         TabIndex        =   6
         Tag             =   "1"
         Top             =   900
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
         MICON           =   "frmInsuranceReports.frx":0060
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
         Caption         =   "Find the total number of people who owned cars that were involved  in accidents in year "
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
         Height          =   840
         Left            =   195
         TabIndex        =   14
         Tag             =   "1"
         Top             =   135
         Width           =   8580
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   540
      TabIndex        =   11
      Top             =   5601
      Width           =   10785
      Begin VB.ComboBox cmbMod 
         Height          =   315
         Left            =   5625
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "1"
         Top             =   855
         Width           =   3315
      End
      Begin ManoharButton.MyButton MyButton2 
         Height          =   375
         Left            =   9045
         TabIndex        =   8
         Tag             =   "1"
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "P&rint"
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
         MICON           =   "frmInsuranceReports.frx":007C
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the model from the drop-down list here"
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
         Left            =   210
         TabIndex        =   17
         Tag             =   "1"
         Top             =   870
         Width           =   4485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Find the number of accidents in which  cars belonging to a specific model were involved."
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
         Height          =   465
         Left            =   195
         TabIndex        =   12
         Tag             =   "1"
         Top             =   135
         Width           =   10335
      End
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   10
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
      Caption         =   "Insurance DataBase - Report Generator"
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
      Width           =   3930
   End
End
Attribute VB_Name = "frmInsuranceRoports"
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
'
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
Case "Person"
    DataEnvironment1.rsCommand1.Close
    DataEnvironment1.Command1
    drsPersonAll.Caption = "Report Generator - Person Table"
    drsPersonAll.WindowState = 2
    drsPersonAll.Show
Case "Car Manufacturer"
    DataEnvironment1.rscmdAllManufacturers.Close
    DataEnvironment1.cmdAllManufacturers
    drsPersonAll.Caption = "Report Generator - Manufacturer Table"
    dsrManufacturerAll.WindowState = 2
    dsrManufacturerAll.Show
Case "Car Makes"
    DataEnvironment1.rscmdAllCarMakes.Close
    DataEnvironment1.cmdAllCarMakes
    drsPersonAll.Caption = "Report Generator - Car Makes Table"
    dsrAllCarMAkes.WindowState = 2
    dsrAllCarMAkes.Show
Case "Car Models"
    DataEnvironment1.rscmdCarModels.Close
    DataEnvironment1.cmdCarModels
    drsPersonAll.Caption = "Report Generator - Car Models Table"
    dsrCarModels.WindowState = 2
    dsrCarModels.Show
Case "Accident"
    DataEnvironment1.rscmdAccident.Close
    DataEnvironment1.cmdAccident
    dsrAccident.Caption = "Report Generator - Car Accident Table"
    dsrAccident.WindowState = 2
    dsrAccident.Show
Case "Car DataBase"
    DataEnvironment1.rscmdCars.Close
    DataEnvironment1.cmdCars
    dsrCars.Caption = "Report Generator - Cars Database Table"
    dsrCars.WindowState = 2
    dsrCars.Show
Case "Owns"
    DataEnvironment1.rscmdOwns.Close
    DataEnvironment1.cmdOwns
    dsrOwns.Caption = "Report Generator - Owns Table"
    dsrOwns.WindowState = 2
    dsrOwns.Show
Case "Participated"
    DataEnvironment1.rscmdParticipated.Close
    DataEnvironment1.cmdParticipated
    dsrParticipated.Caption = "Report Generator - Participated Table"
    dsrParticipated.WindowState = 2
    dsrParticipated.Show

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
'On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Insurance DataBase - Report Generator"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
Dim rsD As New ADODB.Recordset
rsD.Open "select distinct acc_date from accident", Cn, adOpenStatic
If rsD.EOF = True And rsD.BOF = True Then
'
Else

cmbA.CLEAR
cmbA2.CLEAR
Dim Exists As Boolean

rsD.MoveFirst
While Not rsD.EOF
    Exists = False
    For j = 0 To cmbA.ListCount - 1
        If Format(rsD.Fields(0), "yyyy") = cmbA.List(j) Then
            Exists = True
        End If
    Next
    If Exists = False Then
        cmbA.AddItem Format(rsD.Fields(0), "yyyy")
        cmbA2.AddItem Format(rsD.Fields(0), "yyyy")
    End If
    rsD.MoveNext
Wend
cmbA.Text = cmbA.List(0)
cmbA2.Text = cmbA2.List(0)

End If
rsD.Close

rsD.Open "select distinct model from cars where reg_no in (select reg_numb from participated)", Cn, adOpenStatic
If rsD.EOF = True And rsD.BOF = True Then
'
Else
rsD.MoveFirst
cmbMod.CLEAR
While Not rsD.EOF
    cmbMod.AddItem Format(rsD.Fields(0))
    rsD.MoveNext
Wend
cmbMod.Text = cmbMod.List(0)
End If
rsD.Close

Call loadTables
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
frmInsuranceMain.Show
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
    DataEnvironment1.rscmd1_4.Close
    DataEnvironment1.cmd1_4 "%" & Mid(cmbA.Text, 3, 4)
    dsrInsQuery4.Caption = "Report Generator - Person who owns car and involved in accident in year " & cmbA.Text
    dsrInsQuery4.WindowState = 2
    dsrInsQuery4.Show
    
End Sub

Private Sub MyButton2_Click()
On Error Resume Next
    DataEnvironment1.rscmd1_5.Close
    DataEnvironment1.cmd1_5 cmbMod.Text
    dsrInsQuery5.Caption = "Report Generator - Accidents in which cars belonging to the model -' " & cmbMod.Text & " '- were involved"
    dsrInsQuery5.WindowState = 2
    dsrInsQuery5.Show

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
    DataEnvironment1.rscmdEQ2.Close
    DataEnvironment1.cmdEQ2 "%" & Mid(cmbA2.Text, 3, 4)
    dsrInsEQ2.Caption = "Report Generator - People who owns cars and were involved in acidents with their registered car in year " & cmbA2.Text
    dsrInsEQ2.WindowState = 2
    dsrInsEQ2.Show
End Sub

Sub loadTables()
cmbTables.CLEAR
cmbTables.AddItem "Person"
cmbTables.AddItem "Car Manufacturer"
cmbTables.AddItem "Car Makes"
cmbTables.AddItem "Car Models"
cmbTables.AddItem "Car DataBase"
cmbTables.AddItem "Accident"
cmbTables.AddItem "Owns"
cmbTables.AddItem "Participated"
cmbTables.Text = cmbTables.List(0)
End Sub
