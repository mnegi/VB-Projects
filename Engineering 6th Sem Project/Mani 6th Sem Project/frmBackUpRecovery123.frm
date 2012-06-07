VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Begin VB.Form frmBackUpRecovery 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Craeting User Account - *"
   ClientHeight    =   4545
   ClientLeft      =   1740
   ClientTop       =   1995
   ClientWidth     =   8205
   Icon            =   "frmBackUpRecovery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4545
   ScaleWidth      =   8205
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   2700
      TabIndex        =   26
      Top             =   615
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   2325
   End
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
      Left            =   240
      TabIndex        =   3
      Top             =   885
      Width           =   7740
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   2625
         TabIndex        =   10
         Top             =   30
         Visible         =   0   'False
         Width           =   5070
         Begin VB.Label lblDesc 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BackUp Manager 1.0"
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
            TabIndex        =   11
            Tag             =   "1"
            Top             =   180
            Width           =   4965
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   15
         TabIndex        =   4
         Top             =   30
         Width           =   2535
         Begin VB.Label lblThemeSelect 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BackUp Menu"
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
            Left            =   540
            TabIndex        =   0
            Tag             =   "1"
            Top             =   180
            Width           =   1470
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2550
         Left            =   15
         TabIndex        =   5
         Top             =   585
         Width           =   2535
         Begin ManoharButton.MyButton cmdBackUp 
            Height          =   405
            Left            =   105
            TabIndex        =   6
            Tag             =   "1"
            ToolTipText     =   "Start BackUp Manager"
            Top             =   285
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&BackUp Data"
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
            MICON           =   "frmBackUpRecovery.frx":000C
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
            TabIndex        =   7
            Tag             =   "1"
            ToolTipText     =   "Back To Main Menu"
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
            MICON           =   "frmBackUpRecovery.frx":0028
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ManoharButton.MyButton cmbRecover 
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Tag             =   "1"
            ToolTipText     =   "Start Recovery Manager"
            Top             =   870
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&Recover Data"
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
            MICON           =   "frmBackUpRecovery.frx":0044
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
            TabIndex        =   9
            Tag             =   "1"
            ToolTipText     =   "Show Help"
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
            MICON           =   "frmBackUpRecovery.frx":0060
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
      Begin VB.Frame fraRestore 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   2625
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   5070
         Begin VB.ComboBox cmbRType 
            Height          =   315
            ItemData        =   "frmBackUpRecovery.frx":007C
            Left            =   105
            List            =   "frmBackUpRecovery.frx":007E
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "1"
            ToolTipText     =   "Select The Recovery Type."
            Top             =   690
            Width           =   4830
         End
         Begin VB.ComboBox cmbRTables 
            Height          =   315
            ItemData        =   "frmBackUpRecovery.frx":0080
            Left            =   105
            List            =   "frmBackUpRecovery.frx":0082
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "1"
            ToolTipText     =   "Tables Available For Recovery."
            Top             =   1260
            Visible         =   0   'False
            Width           =   4830
         End
         Begin VB.CheckBox chkRestore 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Restore Images/Photos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   19
            ToolTipText     =   "Checking this will import the images as well. This will happen only when images were backed up."
            Top             =   1860
            Width           =   2910
         End
         Begin ManoharButton.MyButton cmdRestore 
            Height          =   405
            Left            =   3465
            TabIndex        =   22
            Tag             =   "1"
            ToolTipText     =   "Restore Database."
            Top             =   1920
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "R&estore"
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
            MICON           =   "frmBackUpRecovery.frx":0084
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox chkDel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delete Existing Data"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   25
            ToolTipText     =   "Checking this will perfoem the update operation if record exists."
            Top             =   2205
            Width           =   2910
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Restore / Recovery Type"
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
            Left            =   75
            TabIndex        =   23
            Tag             =   "1"
            Top             =   150
            Width           =   3210
         End
      End
      Begin VB.Frame fraBackUp 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   2610
         TabIndex        =   12
         Top             =   585
         Visible         =   0   'False
         Width           =   5070
         Begin VB.CheckBox chkBackUp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BackUp Images/Photos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Checking this will save the images as well with data. You can get these back at recovery time."
            Top             =   1830
            Width           =   2910
         End
         Begin VB.ComboBox cmbTables 
            Height          =   315
            ItemData        =   "frmBackUpRecovery.frx":00A0
            Left            =   90
            List            =   "frmBackUpRecovery.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "1"
            ToolTipText     =   "Select the table to backup."
            Top             =   1260
            Visible         =   0   'False
            Width           =   4830
         End
         Begin VB.ComboBox cmbTypes 
            Height          =   315
            ItemData        =   "frmBackUpRecovery.frx":00A4
            Left            =   105
            List            =   "frmBackUpRecovery.frx":00A6
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Tag             =   "1"
            ToolTipText     =   "Select the backup type."
            Top             =   675
            Width           =   4830
         End
         Begin ManoharButton.MyButton MyButton1 
            Height          =   405
            Left            =   3465
            TabIndex        =   17
            Tag             =   "1"
            ToolTipText     =   "Start Backup."
            Top             =   1965
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "&BackUp"
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
            MICON           =   "frmBackUpRecovery.frx":00A8
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
            Caption         =   "Select BackUp type from the list here."
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
            Left            =   75
            TabIndex        =   13
            Tag             =   "1"
            Top             =   150
            Width           =   3810
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
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manohar BackUp - Recovery Manager 1.0"
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
      Width           =   4185
   End
End
Attribute VB_Name = "frmBackUpRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'   CODE BY : MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Private Sub cmbRecover_Click()
fraDesc.Visible = True
fraBackUp.Visible = False
fraRestore.Visible = True
lblDesc.Caption = "Manohar Recovery Manager 1.0"

cmbTypes.CLEAR
cmbRType.CLEAR
cmbTypes.AddItem "Entire DataBase"
cmbRType.AddItem "Entire DataBase"
cmbTypes.AddItem "Individual Tables"
cmbRType.AddItem "Individual Tables"
cmbTypes.Text = cmbTypes.List(0)
cmbRType.Text = cmbRType.List(0)
'
End Sub

Private Sub cmbRType_Click()

If cmbRType.ListIndex = 0 Then

cmdRestore.Enabled = True
cmbRTables.Visible = False

Else

cmbRTables.Visible = True
cmbRTables.CLEAR
File1.Path = App.Path & "\BackUp"
'MsgBox File1.ListCount
Dim rsT As New ADODB.Recordset
rsT.Open "select * from tab", Cn, adOpenStatic
While Not rsT.EOF
For j = 0 To File1.ListCount - 1
If rsT.Fields(0) & ".bkp" = File1.List(j) Then
cmbRTables.AddItem rsT.Fields(0)
End If
Next

rsT.MoveNext
Wend

If cmbRTables.ListCount <> 0 Then
cmbRTables.Text = cmbRTables.List(0)
End If
End If
End Sub

Private Sub cmbTables_Click()
Dim rsB As New ADODB.Recordset
rsB.Open "select * from " & cmbTables.Text, Cn, adOpenStatic
'check if recordset is empty
If rsB.EOF = True And rsB.BOF = True Then
    MyButton1.Enabled = False
        Exit Sub
Else

'if recordset not empty
MyButton1.Enabled = True
'check if the backup folder exists
    If Not MyFile.FolderExists(App.Path & "\BackUp") Then
        MyFile.CreateFolder (App.Path & "\BackUp")
    End If
End If
rsB.Close
End Sub

Private Sub cmbTypes_click()

If cmbTypes.ListIndex = 0 Then

MyButton1.Enabled = True
cmbTables.Visible = False

Else

cmbTables.Visible = True
cmbTables.CLEAR
Dim rsT As New ADODB.Recordset
rsT.Open "select * from tab", Cn, adOpenStatic
While Not rsT.EOF
cmbTables.AddItem rsT.Fields(0)
rsT.MoveNext
Wend

If cmbTables.ListCount <> 0 Then
cmbTables.Text = cmbTables.List(0)
End If
End If

End Sub

Private Sub cmdBackUp_Click()
fraDesc.Visible = True
fraRestore.Visible = False
fraBackUp.Visible = True
lblDesc.Caption = "Manohar BackUp Manager 1.0"

cmbTypes.CLEAR
cmbRType.CLEAR
cmbTypes.AddItem "Entire DataBase"
cmbRType.AddItem "Entire DataBase"
cmbTypes.AddItem "Individual Tables"
cmbRType.AddItem "Individual Tables"
cmbTypes.Text = cmbTypes.List(0)
cmbRType.Text = cmbRType.List(0)

End Sub

Private Sub cmdRestore_Click()
On Error GoTo handle
'Check if entire database is selected
If cmbRType.Text = "Entire DataBase" Then
Dim Arr(30) As String
'insurance database
Arr(0) = "person"
Arr(1) = "manufacturer"
Arr(2) = "carmakes"
Arr(3) = "carmodels"
Arr(4) = "cars"
Arr(5) = "accident"
Arr(6) = "owns"
Arr(7) = "participated"

'order processing database
Arr(8) = "cust"
Arr(9) = "order_tab"
Arr(10) = "item"
Arr(11) = "warehouse"
Arr(12) = "order_item"
Arr(13) = "shipment"

'student enrollment database
Arr(14) = "student"
Arr(15) = "course"
Arr(16) = "text"
Arr(17) = "enroll"
Arr(18) = "book_adoption"

'book dealer
Arr(19) = "author"
Arr(20) = "publisher"
Arr(21) = "category"
Arr(22) = "catalog"
Arr(23) = "order_details"

'bank enterprise database
Arr(24) = "branch"
Arr(25) = "acount"
Arr(26) = "customer"
Arr(27) = "loan"
Arr(28) = "depositor"
Arr(29) = "borrower"

If chkDel.Value = 1 Then
    Call DeleteAllTables
End If

For i = 0 To 30
If MyFile.FileExists(App.Path & "\BackUp\" & Arr(i) & ".bkp") Then
    MsgBox "Table : " & Arr(i) & vbCrLf & "Total Of " & Recover(Arr(i)) & " Records Recovered", vbOKOnly + vbInformation, "Recovered " & Arr(i)
End If
Next


MsgBox "Completed Entire Restore"

'individual tables
Else



'*******************************************
'check if the delete existing data is cheked
' if it is checked delete the records of that
'*******************************************
If chkDel.Value = 1 Then
    Cmd.ActiveConnection = Cn
    Cmd.CommandText = "delete from " & cmbRTables.Text
    Cmd.Execute
End If


MsgBox "Table : " & cmbTables.Text & vbCrLf & "Total Of " & Recover(cmbRTables.Text) & " Records Recovered", vbOKOnly + vbInformation, "Recovered " & cmbTables.Text



'*******************************************
'check if the restore images is checked or not
'*******************************************
If chkRestore.Value = 1 Then
Select Case cmbRTables.Text

Case "PERSON"
If Not MyFile.FolderExists(App.Path & "\Common\Images\MyPhotos") Then
MyFile.CreateFolder (App.Path & "\Common\Images\MyPhotos")
End If


File1.Path = App.Path & "\Common\Images\MyPhotos"
File2.Path = App.Path & "\BackUp\MyPhotos"
Dim exist As Boolean
For i = 0 To File2.ListCount - 1
    For j = 0 To File1.ListCount - 1
        If File2.List(i) = File1.List(j) Then
            exist = True
        Else
            exist = False
        End If
    Next
    If exist = False Then
    MyFile.CopyFile App.Path & "\BackUp\MyPhotos\" & File2.List(i), App.Path & "\Common\Images\MyPhotos\"
    End If
Next


Case "MANUFACTURER"
If MyFile.FolderExists(App.Path & "\BackUp\CarManufacturers") Then
MyFile.DeleteFolder (App.Path & "\BackUp\CarManufacturers")
End If
MyFile.CreateFolder (App.Path & "\BackUp\CarManufacturers")
MyFile.CopyFolder App.Path & "\Common\Images\CarManufacturers", App.Path & "\BackUp\CarManufacturers"


End Select
End If

End If

handle:
If Err.Number <> 0 Then
MsgBox "Error : " & Err.Number & vbCrLf & Err.Description
End If

End Sub
Function Recover(TableName As String) As Integer
On Error GoTo errorHandle


'****************************************
'Open the selected table
'****************************************
Dim Recovered As Integer
Dim RsR As New ADODB.Recordset
RsR.Open "select * from " & TableName, Cn, adOpenStatic, adLockOptimistic

'****************************************
'count the total no of fileds in table
'****************************************
Dim RCount As Integer
RCount = RsR.Fields.Count - 1

'****************************************
'check the field name and its type
'****************************************
'For jp = 0 To RsR.Fields.Count - 1
'MsgBox RsR.Fields(jp).NAME & " : " & RsR.Fields(jp).Type
'Next

Dim Counter As Integer

Dim CmdTxt As String

'****************************************
'Open the backup file (saved with .bkp)
'read that file and do the recovery
'****************************************

Open App.Path & "\BackUp\" & TableName & ".bkp" For Input As #1
Do Until EOF(1)
    Line Input #1, newline
    Counter = 0
    CmdTxt = "insert into " & TableName & " values("
M:
    For k = 1 To Len(newline)
        If Mid(newline, k, 1) = "^" Then
'        MsgBox Counter + 1 & " : " & Mid(newline, 1, k - 1)
        '*************************************
        'check if it is string type
        '200 represents string type
        '*************************************
        If RsR.Fields(Counter).Type = 200 Then
            'check if it is last field
            If Counter = RCount Then
                CmdTxt = CmdTxt & "'" & Mid(newline, 1, k - 1) & "'"
            Else
                CmdTxt = CmdTxt & "'" & Mid(newline, 1, k - 1) & "',"
            End If
        End If
        
        '*************************************
        'check if it is date type
        '135 represents string type
        '*************************************
        If RsR.Fields(Counter).Type = 135 Then
            'check if it is last field
            If Counter = RCount Then
                CmdTxt = CmdTxt & "'" & Format(Mid(newline, 1, k - 1), "dd-mmm-yy") & "'"
            Else
                CmdTxt = CmdTxt & "'" & Format(Mid(newline, 1, k - 1), "dd-mmm-yy") & "',"
            End If
        End If
    '*************************************
    'check if it is number
    '131 represents number type
    '*************************************
    If RsR.Fields(Counter).Type = 131 Then
        'check if it is last field
        If Counter = RCount Then
            CmdTxt = CmdTxt & Mid(newline, 1, k - 1)
        Else
            CmdTxt = CmdTxt & Mid(newline, 1, k - 1) & ","
        End If
    End If
    
    
    Counter = Counter + 1

    newline = Mid(newline, k + 1, Len(newline))
        GoTo M
End If
Next
CmdTxt = CmdTxt & ")"
Cmd.ActiveConnection = Cn
Cmd.CommandText = CmdTxt
Cmd.Execute
Recovered = Recovered + 1
Loop
Close #1

Recover = Recovered
errorHandle:
If Err.Number <> 0 Then
'MsgBox Err.Number & " : " & Err.Description
If Recovered > -1 Then
Recovered = Recovered - 1
End If
Resume Next
End If
End Function
Private Sub cmExit_Click()
frmMain.Show
Unload Me
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
Me.Caption = "Manohar BackUp - Recovery Manager 1.0"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 548, 305, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")


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
frmMain.Show
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
Sub fraoptVisible(X As Integer)
Dim i As Integer
For i = 0 To 2
If i = X Then
fraOpt(i).Visible = True
Else
fraOpt(i).Visible = False
End If
Next
End Sub




Public Function BackUp(rs1 As ADODB.Recordset, Filename As String, FolderName As String, UseFolder As Boolean, Tname As String) As Boolean
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
    .Filename = cmbTables.Text

    .InitDir = App.Path & "\BackUp"
    .DialogTitle = "Save BackUp File"
    .Filter = "Export Files (*.bkp)|*.bkp"
    .FilterIndex = 1
    .DefaultExt = "bkp"
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
    msg = "BackUp operatoin was cancelled" & vbCrLf
    iIndx = MsgBox(msg, vbOKOnly + vbInformation, "BackUp Table Data")
    BackUp = False
    Exit Function
Else
'    On Error GoTo expError
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
            Print #iFileNum, "^"; 'simply a carret string to stay on same line append ;
        Else
'            If iIndx = iNumnberOfFields Then
'                Print #iFileNum, Trim$(CStr(rs1.Fields(iIndx)));
'            Else
                Print #iFileNum, Trim$(CStr(rs1.Fields(iIndx))); "^";
'            End If
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
msg = "BackUp Complete : " & Tname & ".bkp" & vbCrLf
msg = msg & iTotalRecords & " Records Backed Up." & vbCrLf
iIndx = MsgBox(msg, vbOKOnly + vbInformation, "BackUp Table Data")
BackUp = True
Exit Function

expError:

    MsgBox Err.Number & " " & Err.Description
    BackUp = False
End Function

Private Sub MyButton1_Click()
If Not MyFile.FolderExists(App.Path & "\BackUp") Then
    MyFile.CreateFolder (App.Path & "\BackUp")
End If

Dim rs1 As New ADODB.Recordset

If cmbTypes.Text = "Entire DataBase" Then
'CHECK IF USER SELECTS THE ENTIRE DATABSE BACKUP
Dim rsT As New ADODB.Recordset

If MyFile.FolderExists(App.Path & "\BackUp\Entire DB") Then
MyFile.DeleteFolder (App.Path & "\BackUp\Entire DB")
Else
MyFile.CreateFolder (App.Path & "\BackUp\Entire DB")
End If

rsT.Open "select * from tab", Cn, adOpenStatic
While Not rsT.EOF
cmbTables.AddItem rsT.Fields(0)
rsT.MoveNext
Wend
Dim Counter As Integer
For i = 0 To cmbTables.ListCount - 1
    sqlStr = "select * from " & cmbTables.List(i)
'    rs1.CursorLocation = adUseClient
    rs1.Open sqlStr, Cn
    If Not (rs1.EOF = True And rs1.BOF = True) Then
    Call BackUp(rs1, "", App.Path & "\BackUp\Entire DB\" & cmbTables.List(i) & ".bkp", True, cmbTables.List(i))
    Counter = Counter + 1
    rs1.Close
    Else
    rs1.Close
    End If
Next
MsgBox "TOTAL NUMBER OF TABLES BACKED UP ARE : " & Counter
'If chkBackUp.Value = 1 Then

Else


sqlStr = "select * from " & cmbTables.Text
rs1.CursorLocation = adUseClient
rs1.Open sqlStr, Cn
Call BackUp(rs1, cmbTables.Text & ".bkp", "", False, cmbTables.Text)
If chkBackUp.Value = 1 Then
Select Case cmbTables.Text

Case "PERSON"
'Case "STUDENT"
'Case "CUST"
'Case "AUTHOR"
'Case "CUSTOMER"

If MyFile.FolderExists(App.Path & "\BackUp\MyPhotos") Then
MyFile.DeleteFolder (App.Path & "\BackUp\MyPhotos")
End If

MyFile.CreateFolder (App.Path & "\BackUp\MyPhotos")
MyFile.CopyFolder App.Path & "\Common\Images\MyPhotos", App.Path & "\BackUp\MyPhotos"

Case "MANUFACTURER"
If MyFile.FolderExists(App.Path & "\BackUp\CarManufacturers") Then
MyFile.DeleteFolder (App.Path & "\BackUp\CarManufacturers")
End If
MyFile.CreateFolder (App.Path & "\BackUp\CarManufacturers")
MyFile.CopyFolder App.Path & "\Common\Images\CarManufacturers", App.Path & "\BackUp\CarManufacturers"


End Select
End If
rs1.Close

End If

End Sub
