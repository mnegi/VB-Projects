VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmOrderItem 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Craeting User Account - *"
   ClientHeight    =   4545
   ClientLeft      =   1740
   ClientTop       =   1995
   ClientWidth     =   8205
   Icon            =   "frmOrderItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2460
      Left            =   255
      TabIndex        =   13
      Top             =   840
      Width           =   5805
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         Left            =   2265
         TabIndex        =   19
         Tag             =   "1"
         Top             =   375
         Width           =   3435
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   2265
         TabIndex        =   18
         Tag             =   "1"
         Text            =   "Combo1"
         Top             =   1095
         Width           =   3435
      End
      Begin XPTEXTBOX.text txtQty 
         Height          =   420
         Left            =   2265
         TabIndex        =   14
         Tag             =   "1"
         ToolTipText     =   "LOCATION OF ACCIDENT"
         Top             =   1815
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   741
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Number"
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
         Left            =   105
         TabIndex        =   17
         Tag             =   "1"
         Top             =   1095
         Width           =   1320
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number"
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
         Left            =   105
         TabIndex        =   16
         Tag             =   "1"
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label lblMidName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   120
         TabIndex        =   15
         Tag             =   "1"
         Top             =   1830
         Width           =   840
      End
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   435
      Left            =   1395
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   375
      Index           =   0
      Left            =   6180
      TabIndex        =   2
      Tag             =   "1"
      Top             =   915
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add New"
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
      MICON           =   "frmOrderItem.frx":000C
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
      Height          =   375
      Index           =   1
      Left            =   6150
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Save"
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
      MICON           =   "frmOrderItem.frx":0028
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
      Height          =   375
      Index           =   3
      Left            =   6150
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1755
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Edit"
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
      MICON           =   "frmOrderItem.frx":0044
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
      Height          =   375
      Index           =   4
      Left            =   6150
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2175
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Update"
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
      MICON           =   "frmOrderItem.frx":0060
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
      Height          =   375
      Index           =   5
      Left            =   6150
      TabIndex        =   6
      Tag             =   "1"
      Top             =   2595
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderItem.frx":007C
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
      Height          =   375
      Index           =   2
      Left            =   6150
      TabIndex        =   7
      Tag             =   "1"
      Top             =   3030
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "frmOrderItem.frx":0098
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
      Height          =   375
      Index           =   6
      Left            =   6180
      TabIndex        =   8
      Tag             =   "1"
      Top             =   3465
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Main"
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
      MICON           =   "frmOrderItem.frx":00B4
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
      Height          =   375
      Index           =   7
      Left            =   165
      TabIndex        =   9
      Tag             =   "1"
      Top             =   3465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&First"
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
      MICON           =   "frmOrderItem.frx":00D0
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
      Height          =   375
      Index           =   8
      Left            =   1700
      TabIndex        =   10
      Tag             =   "1"
      Top             =   3465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Previous"
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
      MICON           =   "frmOrderItem.frx":00EC
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
      Height          =   375
      Index           =   9
      Left            =   3235
      TabIndex        =   11
      Tag             =   "1"
      Top             =   3465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Next"
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
      MICON           =   "frmOrderItem.frx":0108
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
      Height          =   375
      Index           =   10
      Left            =   4770
      TabIndex        =   12
      Tag             =   "1"
      Top             =   3465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Last"
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
      MICON           =   "frmOrderItem.frx":0124
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Item Database"
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
      TabIndex        =   0
      Tag             =   "1"
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "frmOrderItem"
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

Dim Saved As Boolean
Dim loaded As Boolean
Dim rs1 As New ADODB.Recordset

Private Sub Btns_Click(Index As Integer)
''On Error GoTo q
Select Case Index

Case 0
'add new record
Call CLEAR
Saved = False
LockTheControls (False)
Dim RsC As New ADODB.Recordset
RsC.Open "select * from order_tab", Cn, adOpenStatic
cmbOrder.CLEAR
If RsC.EOF = True And RsC.BOF = True Then
MsgBox "You should have some orders"
Call btnEnable("10110111111")
Exit Sub

Else
RsC.MoveFirst
While Not RsC.EOF
cmbOrder.AddItem RsC.Fields(0)
RsC.MoveNext
Wend
RsC.Close

RsC.Open "select * from item", Cn, adOpenStatic
cmbItem.CLEAR
If RsC.EOF = True And RsC.BOF = True Then
MsgBox "You should have some items"
Call btnEnable("10110111111")
Exit Sub

Else
RsC.MoveFirst
While Not RsC.EOF
cmbItem.AddItem RsC.Fields(0)
RsC.MoveNext
Wend
RsC.Close
End If



Call LockTheControls(False)
Call btnEnable("01000100000")
End If

Case 1
'save record
On Error GoTo HandleErr

Cmd.ActiveConnection = Cn
Cmd.CommandText = "insert into order_item values(" & cmbOrder.Text & "," & cmbItem.Text & "," & txtQty.Text & ")"
MsgBox Cmd.CommandText
Cmd.Execute
Saved = True
Call btnEnable("10110111111")
LockTheControls (True)
Call Load_Records
rs1.Requery


HandleErr:
If Err.Number <> 0 Then
Call Handle_Error(Err.Number, "Error : " & Err.Number, Err.Description, "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
Exit Sub
End If

Case 2
'delete
Call Handle_Error("Confirm Delete", "Deleting Current record", "This will delete the currently displayed record. Are you sure to delete ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If Not MsgBOx_R_Value Then
Exit Sub
Else
Cmd.ActiveConnection = Cn
Dim Cstring As String
Cstring = "delete from order_item where i_order=" & cmbOrder.Text & " it=" & cmbItem.Text
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute
rs1.Requery
End If
Call CLEAR
Call Load_Records
Call LockTheControls(True)


Case 3
'edit
LockTheControls (False)
Saved = False
btnEnable ("0000110000")
cmbOrder.Locked = True
cmbOrder.BackColor = vbWhite
cmbItem.Locked = True
cmbItem.BackColor = vbWhite

Case 4
'Update
Cmd.ActiveConnection = Cn

Cstring = "update order_item set qut =" & txtQty.Text & " where i_order=" & cmbOrder.Text & " it=" & cmbItem.Text
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute
Saved = True
rs1.Requery
Call Load_Records
Call LockTheControls(True)

Case 5
''cancel
Saved = True
LockTheControls (True)
Call btnEnable("10110111111")
Call Load_Records

Case 6
Call ExitPrg
Case 7

If Not rs1.BOF Then
rs1.MoveFirst
Call DISPLAY(rs1)
If rs1.BOF = True Then
Btns(7).Enabled = False
End If
End If

Case 8
If Not rs1.BOF Then
rs1.MovePrevious
If rs1.BOF Then rs1.MoveFirst
End If
Call DISPLAY(rs1)

Case 9

If Not rs1.EOF Then
rs1.MoveNext
If rs1.EOF Then rs1.MoveLast

End If
Call DISPLAY(rs1)

Case 10
If Not rs1.EOF Then
rs1.MoveLast
End If
Call DISPLAY(rs1)

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
Call Apply_Theme(Me, 1)
If Saved = True Then
Call LockTheControls(True)
loaded = True
Call Load_Records

End If


q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Private Sub Form_DblClick()
Me.Left = 1740
Me.Top = 1995
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

rs1.Open "select * from ORDER_item", Cn, adOpenDynamic
Saved = True
loaded = False
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")

Me.Caption = "Order Item DataBase"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 548, 305, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")


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
Me.Left = 1740
Me.Top = 1995
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
Call ExitPrg
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
Sub CLEAR()
cmbOrder.CLEAR
cmbItem.CLEAR
txtQty.Text = ""
End Sub
Sub DISPLAY(DRS As ADODB.Recordset)
cmbOrder.CLEAR
If Not IsNull(DRS.Fields(0)) Then
cmbOrder.AddItem DRS.Fields(0)
cmbOrder.Text = cmbOrder.List(0)
Else
'
End If
cmbItem.CLEAR
If Not IsNull(DRS.Fields(1)) Then
cmbItem.AddItem DRS.Fields(1)
cmbItem.Text = cmbItem.List(0)
Else
'
End If

If Not IsNull(DRS.Fields(2)) Then
txtQty.Text = DRS.Fields(2)
Else
txtQty.Text = ""
End If

End Sub

Sub LockTheControls(btn As Boolean)

With Me
For iindex = 0 To .Controls.Count - 1
If .Controls(iindex).Tag = "1" Then
If (btn) Then
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox Or TypeOf .Controls(iindex) Is RichTextBox) Then

                    .Controls(iindex).Locked = True
                    .Controls(iindex).BackColor = vbWhite

End If
If (TypeOf .Controls(iindex) Is MaskEdBox Or TypeOf .Controls(iindex) Is ImageCombo) Then

                    .Controls(iindex).Enabled = False
                    .Controls(iindex).BackColor = vbWhite

End If

Else
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox Or TypeOf .Controls(iindex) Is RichTextBox) Then
                    .Controls(iindex).Locked = False
                    .Controls(iindex).BackColor = TextBackcolor


End If
If (TypeOf .Controls(iindex) Is MaskEdBox Or TypeOf .Controls(iindex) Is ImageCombo) Then

                    .Controls(iindex).Enabled = True
                    .Controls(iindex).BackColor = TextBackcolor

End If
End If

End If
Next
End With
End Sub

Sub btnEnable(btn As String)
'MsgBox Len(btn)
For i = 1 To Len(btn)
If Mid$(btn, i, 1) = "1" Then
Btns(i - 1).Enabled = True
Else
Btns(i - 1).Enabled = False
End If
Next
End Sub

Sub Load_Records()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from order_item"
Set rs = Cmd.Execute
If rs.BOF = True And rs.EOF = True Then
Call CLEAR
Call btnEnable("10000010000")
Else
rs.MoveFirst
Call DISPLAY(rs)
Call btnEnable("10110011111")
End If
End Sub

Sub ExitPrg()
'On Error Resume Next
If Saved Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmOrderMain.Show
rs1.Close
Else

Call Handle_Error("Confirm", "Record Not Saved", "Current record not yet saved. Do you want to exit ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If MsgBOx_R_Value Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmOrderMain.Show
Else
Exit Sub
End If
End If
End Sub

Private Sub txtodate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtLocation.SetFocus
Else
Call DATEVALID(KeyAscii)
End If
End Sub

Private Sub txtLocation_Change()
txtLocation.Text = StrConv(txtLocation.Text, vbProperCase)
SendKeys "{END}"
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'
Else
Call ADDVALID(KeyAscii)
End If
End Sub

Private Sub txtOno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtOdate.SetFocus
Else
Call NUMVALID(KeyAscii)
End If
End Sub
