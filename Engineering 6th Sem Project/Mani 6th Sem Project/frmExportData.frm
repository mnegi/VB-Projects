VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
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
   Icon            =   "frmExportData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4545
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6600
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   270
      Left            =   1395
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   885
      Width           =   7740
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   2625
         TabIndex        =   15
         Top             =   45
         Visible         =   0   'False
         Width           =   5070
         Begin VB.Label lblDesc 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Data Format"
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
            Left            =   60
            TabIndex        =   16
            Tag             =   "1"
            Top             =   165
            Width           =   4965
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   15
         TabIndex        =   4
         Top             =   45
         Width           =   2535
         Begin VB.Label lblThemeSelect 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Menu"
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
            Left            =   60
            TabIndex        =   0
            Tag             =   "1"
            Top             =   180
            Width           =   2430
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2550
         Left            =   15
         TabIndex        =   9
         Top             =   600
         Width           =   2535
         Begin ManoharButton.MyButton cmdExpFormats 
            Height          =   405
            Left            =   90
            TabIndex        =   10
            Tag             =   "1"
            ToolTipText     =   "Enables to select among the various export formats"
            Top             =   255
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&Export Formats"
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
            MICON           =   "frmExportData.frx":000C
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
            Left            =   90
            TabIndex        =   11
            Tag             =   "1"
            ToolTipText     =   "Back to main"
            Top             =   2040
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&Main Menu"
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
            MICON           =   "frmExportData.frx":0028
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ManoharButton.MyButton cmdExport 
            Height          =   405
            Left            =   105
            TabIndex        =   12
            Tag             =   "1"
            ToolTipText     =   "Starts the export"
            Top             =   840
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Export &What ?"
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
            MICON           =   "frmExportData.frx":0044
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ManoharButton.MyButton cmdHelp 
            Height          =   405
            Left            =   120
            TabIndex        =   13
            Tag             =   "1"
            ToolTipText     =   "Get the help on this topic"
            Top             =   1425
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&Help"
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
            MICON           =   "frmExportData.frx":0060
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame fraExport 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   2625
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   5070
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmExportData.frx":007C
            Left            =   75
            List            =   "frmExportData.frx":0089
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Tag             =   "1"
            Top             =   495
            Width           =   2880
         End
         Begin VB.Frame fraOpt 
            BackColor       =   &H00FFFFFF&
            Height          =   1515
            Index           =   2
            Left            =   30
            TabIndex        =   20
            Top             =   990
            Visible         =   0   'False
            Width           =   5010
            Begin XPTEXTBOX.text txtQuery 
               Height          =   435
               Left            =   45
               TabIndex        =   22
               Tag             =   "1"
               ToolTipText     =   "Emter your SQL Command here"
               Top             =   510
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   767
               FontName        =   "MS Serif"
               FontSize        =   9.75
               FontBold        =   -1  'True
               LineColor       =   11643476
               Text            =   ""
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin ManoharButton.MyButton cmdQuery 
               Height          =   405
               Left            =   1365
               TabIndex        =   23
               Tag             =   "1"
               ToolTipText     =   "Export the query output"
               Top             =   1020
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   714
               BTYPE           =   3
               TX              =   "Export Query O/P"
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
               MICON           =   "frmExportData.frx":00B7
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblShow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enter the SQL Query"
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
               Left            =   60
               TabIndex        =   21
               Tag             =   "1"
               Top             =   165
               Width           =   2145
            End
         End
         Begin VB.Frame fraOpt 
            BackColor       =   &H00FFFFFF&
            Height          =   1515
            Index           =   1
            Left            =   30
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   5010
            Begin VB.ComboBox cmbTable 
               Height          =   315
               ItemData        =   "frmExportData.frx":00D3
               Left            =   930
               List            =   "frmExportData.frx":00E0
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Tag             =   "1"
               Top             =   180
               Width           =   3975
            End
            Begin ManoharButton.MyButton MyButton1 
               Height          =   405
               Left            =   1515
               TabIndex        =   28
               Tag             =   "1"
               ToolTipText     =   "Export Table Data"
               Top             =   915
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   714
               BTYPE           =   3
               TX              =   "Export &Table"
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
               MICON           =   "frmExportData.frx":010E
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
               Caption         =   "Table"
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
               Left            =   60
               TabIndex        =   29
               Tag             =   "1"
               Top             =   165
               Width           =   570
            End
         End
         Begin VB.Frame fraOpt 
            BackColor       =   &H00FFFFFF&
            Height          =   1515
            Index           =   0
            Left            =   30
            TabIndex        =   24
            Top             =   975
            Visible         =   0   'False
            Width           =   5010
            Begin ManoharButton.MyButton cndExportDB 
               Height          =   405
               Left            =   1350
               TabIndex        =   25
               Tag             =   "1"
               ToolTipText     =   "Export All data from this connection or user"
               Top             =   825
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   714
               BTYPE           =   3
               TX              =   "&Export Databse"
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
               MICON           =   "frmExportData.frx":012A
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Export Entire DataBase"
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
               Left            =   60
               TabIndex        =   26
               Tag             =   "1"
               Top             =   165
               Width           =   2415
            End
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "What do you want to export ?"
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
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Tag             =   "1"
            Top             =   150
            Width           =   4785
         End
      End
      Begin VB.Frame fraExpFormats 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   2625
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   5070
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Export to Microsoft Excel Format."
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
            Height          =   360
            Left            =   150
            TabIndex        =   8
            Top             =   1830
            Width           =   4065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Export to HTML format."
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
            Height          =   360
            Left            =   150
            TabIndex        =   7
            Top             =   1357
            Width           =   4065
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Export to CSV format."
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
            Height          =   360
            Left            =   135
            TabIndex        =   6
            Top             =   900
            Width           =   4065
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select one of the following available export formats"
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
            Height          =   750
            Left            =   90
            TabIndex        =   14
            Tag             =   "1"
            Top             =   150
            Width           =   4785
         End
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
      Caption         =   "Manohar Data Export Manager 1.0"
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
      TabIndex        =   1
      Tag             =   "1"
      Top             =   60
      Width           =   3435
   End
End
Attribute VB_Name = "frmExportData"
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

Option Explicit
Dim cn1 As ADODB.Connection
Dim rs1 As ADODB.Recordset
Dim connStr As String
Dim sqlStr As String
Dim objExcel As Object
Dim objTemp As Object

Private Sub cmbType_Click()
Call fraoptVisible(cmbType.ListIndex)

Dim rsD As New ADODB.Recordset
rsD.Open "select * from tab", cn1, adOpenStatic
cmbTable.CLEAR
While Not rsD.EOF
cmbTable.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
cmbTable.Text = cmbTable.List(0)
rsD.Close


End Sub

Private Sub cmdExpFormats_Click()
lblDesc.Caption = "Export Data Format"
fraDesc.Visible = True
fraExport.Visible = False
fraExpFormats.Visible = True
End Sub

Private Sub cmdExport_Click()
lblDesc.Caption = "Export Data From ?"
fraDesc.Visible = True
fraExpFormats.Visible = False
fraExport.Visible = True
cmbType.Text = cmbType.List(2)


End Sub

Private Sub cmdQuery_Click()

If Option1.Value = True Then
sqlStr = txtQuery.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1
'MsgBox rs1.RecordCount
If rs1.EOF = True And rs1.BOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"
Else
Call CSV(rs1, "Export.csv", "", False)
End If
rs1.Close
MsgBox "DONE"

End If


If Option2.Value = True Then
sqlStr = txtQuery.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1
'MsgBox rs1.RecordCount
If rs1.EOF = True And rs1.BOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"
Else
Call HTML(rs1, "Export.htm", "", False)
End If
rs1.Close
MsgBox "DONE"

End If

If Option3.Value = True Then
sqlStr = txtQuery.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1
'MsgBox rs1.RecordCount
If rs1.EOF = True And rs1.BOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"
Else
Call EXCEL(rs1)
End If
rs1.Close

End If

End Sub

Private Sub cmExit_Click()
frmMain.Show
Unload Me
End Sub

Private Sub cndExportDB_Click()
If MyFile.FolderExists(App.Path & "\Exported Data\Entire DB") Then
'
Else
MyFile.CreateFolder (App.Path & "\Exported Data\Entire DB")
End If
Dim rsD As New ADODB.Recordset
rsD.Open "select * from tab", cn1, adOpenStatic

If Option1.Value = True Then

While Not rsD.EOF

    sqlStr = "select * from " & rsD.Fields(0)
    rs1.CursorLocation = adUseClient
    rs1.Open sqlStr, cn1
    If Not (rs1.EOF = True And rs1.BOF = True) Then
    Call CSV(rs1, "", App.Path & "\Exported Data\Entire DB\" & rsD.Fields(0) & ".csv", True)
    End If
    rs1.Close

    rsD.MoveNext
Wend

Else

If Option2.Value = True Then
While Not rsD.EOF

    sqlStr = "select * from " & rsD.Fields(0)
    rs1.CursorLocation = adUseClient
    rs1.Open sqlStr, cn1
    If Not (rs1.EOF = True And rs1.BOF = True) Then
    Call HTML(rs1, "", App.Path & "\Exported Data\Entire DB\" & rsD.Fields(0) & ".htm", True)
    End If
    rs1.Close

    rsD.MoveNext
Wend


Else
If Option3.Value = True Then
MsgBox "this option is not available for this type of export"
End If
End If
End If
MsgBox "DONE"

rsD.Close
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
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Manohar Data Export Manager 1.0"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 548, 305, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
Option1.Value = True

Set cn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

connStr = "Provider=MSDAORA.1;User ID =" & GetSetting(App.title, "Users", "Name") & ";Password=" & GetSetting(App.title, "Users", "Password")
cn1.Open connStr

If MyFile.FolderExists(App.Path & "\Exported Data") Then
'
Else
MyFile.CreateFolder (App.Path & "\Exported Data")
End If

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
frmMain.Show
Unload Me
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
Sub fraoptVisible(x As Integer)
Dim i As Integer
For i = 0 To 2
If i = x Then
fraOpt(i).Visible = True
Else
fraOpt(i).Visible = False
End If
Next
End Sub

Public Function CSV(adorecordset As ADODB.Recordset, Filename As String, FolderName As String, UseFolder As Boolean) As Boolean
Dim iTotalRecords As Integer
Dim sFileToExport As String
Dim iFileNum As Integer
Dim msg As String
Dim iIndx As Integer
Dim iNumnberOfFields As Integer
If UseFolder = False Then

'Screen.MousePointer = vbDefault
On Error Resume Next
With cd1
    .CancelError = True
    .Filename = Filename
      
    .InitDir = App.Path & "\Exported Data"
    .DialogTitle = "Save Comma Delimited Export File"
    .Filter = "Export Files (*.csv)|*.csv"
    .FilterIndex = 1
    .DefaultExt = "csv"
    .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
    .ShowSave
End With
sFileToExport = cd1.Filename

Else
    sFileToExport = FolderName
End If


'*****************************
'--- user cancels operation
'*****************************
If (Err = 32755) Then 'operation cancelled
    Beep
    msg = "Export operatoin was cancelled" & vbCrLf
    iIndx = MsgBox(msg, vbOKOnly + vbInformation, "Comma Delimited Export File")
    CSV = False
    Exit Function
Else
    On Error GoTo expError
End If

'****************************
'save the data now
'****************************

iTotalRecords = 0

iFileNum = FreeFile()
Open sFileToExport For Output As #iFileNum 'open file for output

'****************************
'stream out the data
'****************************

iNumnberOfFields = rs1.Fields.Count - 1
rs1.MoveFirst
Do Until rs1.EOF
    iTotalRecords = iTotalRecords + 1
    For iIndx = 0 To iNumnberOfFields
        If (IsNull(rs1.Fields(iIndx))) Then
            Print #iFileNum, ","; 'simply a comma string
        Else
            If iIndx = iNumnberOfFields Then
                Print #iFileNum, Trim$(CStr(rs1.Fields(iIndx)));
            Else
                Print #iFileNum, Trim$(CStr(rs1.Fields(iIndx))); ",";
            End If
        End If
    Next
    
    Print #iFileNum,
    rs1.MoveNext
    DoEvents
    
Loop

'****************************
'close ifilenum
'****************************

Close #iFileNum
Beep
msg = "Export File " & sFileToExport & vbCrLf
msg = msg & iTotalRecords & " records written to disk." & vbCrLf
iIndx = MsgBox(msg, vbOKOnly + vbInformation, "Comma Delimited File")
CSV = True
Exit Function

expError:

    MsgBox Err.Number & " " & Err.Description
    CSV = False
End Function
Function CheckQuery(Query As String) As Boolean
On Error GoTo wrong
Dim c1 As New ADODB.Command
c1.ActiveConnection = cn1
c1.CommandText = Query
c1.Execute

CheckQuery = True
Exit Function
wrong:
CheckQuery = False
Exit Function

End Function

Private Sub MyButton1_Click()

If Option1.Value = True Then

sqlStr = "select * from " & cmbTable.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1

If rs1.EOF = True And rs1.BOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"

Else
Call CSV(rs1, cmbTable.Text & ".csv", "", False)
End If
rs1.Close
MsgBox "DONE"



Else

If Option2.Value = True Then

sqlStr = "select * from " & cmbTable.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1

If rs1.EOF = True And rs1.BOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"

Else
Call HTML(rs1, cmbTable.Text & ".HTM", "", False)
End If

rs1.Close
MsgBox "DONE"



Else
sqlStr = "select * from " & cmbTable.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, cn1
If rs1.BOF = True And rs1.EOF = True Then
MsgBox "Empty Table " & vbCrLf & "Nothing to export"
Else
Call EXCEL(rs1)
End If
rs1.Close


End If
End If
End Sub

Private Sub txtQuery_Change()
Dim x As Boolean
x = CheckQuery(txtQuery.Text)
If x = True Then
cmdQuery.Enabled = True
Else
cmdQuery.Enabled = False
End If

End Sub

Public Function HTML(adorecordset As ADODB.Recordset, Filename As String, FolderName As String, UseFolder As Boolean) As Boolean
Dim iTotalRecords As Integer
Dim sFileToExport As String
Dim iFileNum As Integer
Dim msg As String
Dim iIndx As Integer
Dim innerLoop As Integer
Dim iNumnberOfFields As Integer
If UseFolder = False Then

'Screen.MousePointer = vbDefault
On Error Resume Next
With cd1
    .CancelError = True
    .Filename = Filename
      
    .InitDir = App.Path & "\Exported Data"
    .DialogTitle = "Save HTML Export File"
    .Filter = "Export Files (*.htm)|*.htm"
    .FilterIndex = 1
    .DefaultExt = "htm"
    .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
    .ShowSave
End With
sFileToExport = cd1.Filename

Else
    sFileToExport = FolderName
End If


'*****************************
'--- user cancels operation
'*****************************
If (Err = 32755) Then 'operation cancelled
    Beep
    msg = "Export operatoin was cancelled" & vbCrLf
    iIndx = MsgBox(msg, vbOKOnly + vbInformation, "Comma Delimited Export File")
    HTML = False
    Exit Function
Else
    On Error GoTo expError
End If

'****************************
'save the data now
'****************************

iTotalRecords = 0

iFileNum = FreeFile()
Open sFileToExport For Output As #iFileNum 'open file for output

'****************************
'stream out the data
'****************************

iNumnberOfFields = rs1.Fields.Count - 1
rs1.MoveFirst
Print #iFileNum, "<HTML><HEAD><TITLE>HTML Data Export</TITLE></HEAD>"
Print #iFileNum, "<BODY BGCOLOR=""FF00CC"">"

Print #iFileNum, "<TABLE BGCOLOR=""ABDAFF"" WIDTH=""100%"" ALIGN=""CENTER"">"
Print #iFileNum, "<TR><TD>"
Print #iFileNum, "<FONT FACE=TIMES NEW ROMAN SIZE=5 ALIGN=CENTER><B>MANOHAR DATA EXPORT WIZARD 1.0 - HTML FORMAT </B></FONT></TD></TR>"
Print #iFileNum, "</TABLE>"
Print #iFileNum, "<BR><BR>"
Print #iFileNum, "<TABLE BGCOLOR=""00AAFF"">"
Print #iFileNum, "<TR><TD>"
Print #iFileNum, "<FONT FACE=TIMES NEW ROMAN SIZE=5 ALIGN=""LEFT""><B></B></FONT></TD></TR>"
Print #iFileNum, "<TR>"
Print #iFileNum, "<FONT FACE=TIMES NEW ROMAN SIZE+=3>"
For iIndx = 0 To rs1.Fields.Count - 1
    Print #iFileNum, "<TD BGCOLOR=""CCCCC0"">"
    Print #iFileNum, "<B> &nbsp"; rs1.Fields(iIndx).Name; "&nbsp </B>"
    Print #iFileNum, "</TD>"
Next
    
Print #iFileNum, "</TR>"

With rs1
    .MoveFirst
    While Not .EOF
    Print #iFileNum, "<TR>"
    For innerLoop = 0 To .Fields.Count - 1
        Print #iFileNum, "<TD BGCOLOR=""CCCCC0"">"
        Print #iFileNum, "&nbsp"; .Fields(innerLoop); "&nbsp"
        Print #iFileNum, "</TD>"
    Next
    Print #iFileNum, "</TR>"
    .MoveNext
    Wend
End With
Print #iFileNum, "</FONT>"
Print #iFileNum, "</TABLE></BODY></HTML>"
'****************************
'close ifilenum
'****************************
Close #iFileNum
Beep
HTML = True
Exit Function

expError:

    MsgBox Err.Number & " " & Err.Description
   HTML = False
End Function
Public Sub EXCEL(adorecordset As ADODB.Recordset)
Dim iIndx As Integer
Dim iRowIndex As Integer
Dim iColIndex As Integer
Dim iRecordCount As Integer
Dim iFieldCount As Integer
Dim sMessage As String
Dim avRows As Variant
Dim excelVersion As Integer

'READ ALL THE DATA INTO THE ARRAY avRows
avRows = adorecordset.GetRows()

'DETERMINE HOW MANY FIELDS AND RECORDS
iRecordCount = UBound(avRows, 2) + 1
iFieldCount = UBound(avRows, 1) + 1

'CREATE REFERNECE VARIABLE FOR SPREADSHEET
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

'we need this line to ensure Excel remains visible if we switch to the active sheet
Set objTemp = objExcel

excelVersion = Val(objExcel.Application.Version)
If (excelVersion >= 8) Then
    Set objExcel = objExcel.ActiveSheet
End If

'place the name of the fields as the column headers
iRowIndex = 1
iColIndex = 1
For iColIndex = 1 To iFieldCount
With objExcel.Cells(iRowIndex, iColIndex)
    .Value = adorecordset.Fields(iColIndex - 1).Name
    With .Font
        .Name = "Ariel"
        .Bold = True
        .Size = 9
    End With
End With
Next

'memory management
With objExcel
    For iRowIndex = 2 To iRecordCount + 1
    For iColIndex = 1 To iFieldCount
        .Cells(iRowIndex, iColIndex).Value = avRows(iColIndex - 1, iRowIndex - 2)
    Next
    Next
End With
    
objExcel.Cells(1, 1).CurrentRegion.EntireColumn.AutoFit

End Sub

