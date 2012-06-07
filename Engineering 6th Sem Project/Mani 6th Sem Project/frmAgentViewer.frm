VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmAgentViewer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmAgentViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd2 
      Left            =   6210
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouse Coordinates"
      Enabled         =   0   'False
      Height          =   840
      Left            =   300
      TabIndex        =   23
      Tag             =   "1"
      Top             =   6210
      Width           =   2790
      Begin VB.TextBox txtMX 
         Height          =   345
         Left            =   540
         TabIndex        =   25
         Tag             =   "1"
         Top             =   375
         Width           =   780
      End
      Begin VB.TextBox txtMY 
         Height          =   345
         Left            =   1890
         TabIndex        =   24
         Tag             =   "1"
         Top             =   375
         Width           =   780
      End
      Begin VB.Label CharPosnLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
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
         Index           =   3
         Left            =   90
         TabIndex        =   27
         Tag             =   "1"
         Top             =   390
         Width           =   285
      End
      Begin VB.Label CharPosnLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y :"
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
         Index           =   2
         Left            =   1455
         TabIndex        =   26
         Tag             =   "1"
         Top             =   375
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Character Coordinates"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   300
      TabIndex        =   14
      Tag             =   "1"
      Top             =   7155
      Width           =   2790
      Begin VB.TextBox txtCY 
         Height          =   345
         Left            =   1890
         TabIndex        =   17
         Tag             =   "1"
         Top             =   345
         Width           =   780
      End
      Begin VB.TextBox txtCX 
         Height          =   345
         Left            =   540
         TabIndex        =   16
         Tag             =   "1"
         Top             =   345
         Width           =   780
      End
      Begin ManoharButton.MyButton Command1 
         Height          =   405
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Tag             =   "1"
         Top             =   810
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Move"
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
         MICON           =   "frmAgentViewer.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label CharPosnLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y :"
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
         Index           =   1
         Left            =   1455
         TabIndex        =   19
         Tag             =   "1"
         Top             =   345
         Width           =   285
      End
      Begin VB.Label CharPosnLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X :"
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
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Tag             =   "1"
         Top             =   345
         Width           =   285
      End
   End
   Begin VB.Frame SpeechOut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speech &Output"
      Enabled         =   0   'False
      Height          =   2610
      Left            =   3330
      TabIndex        =   8
      Tag             =   "1"
      Top             =   5880
      Width           =   8505
      Begin VB.CheckBox BalloonStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A&uto hide"
         Height          =   300
         Index           =   1
         Left            =   5850
         TabIndex        =   13
         Tag             =   "1"
         Top             =   660
         Width           =   2505
      End
      Begin VB.CheckBox BalloonStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto &pace"
         Height          =   300
         Index           =   2
         Left            =   5850
         TabIndex        =   12
         Tag             =   "1"
         Top             =   1005
         Width           =   2505
      End
      Begin VB.CheckBox BalloonStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Si&ze to text"
         Height          =   300
         Index           =   3
         Left            =   5850
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1350
         Width           =   2505
      End
      Begin VB.CheckBox BalloonStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display &word balloon"
         Height          =   285
         Index           =   0
         Left            =   5535
         TabIndex        =   10
         Tag             =   "1"
         Top             =   285
         Width           =   2880
      End
      Begin VB.TextBox SpeakText 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "1"
         Top             =   255
         Width           =   5340
      End
      Begin ManoharButton.MyButton Command1 
         Height          =   390
         Index           =   2
         Left            =   6360
         TabIndex        =   20
         Tag             =   "1"
         Top             =   2040
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "&Speak"
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
         MICON           =   "frmAgentViewer.frx":0028
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
   Begin VB.Frame AnimationFrame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Animations"
      Enabled         =   0   'False
      Height          =   4695
      Left            =   3330
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1065
      Width           =   8460
      Begin VB.CheckBox OutputStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play sound &effects"
         Height          =   330
         Index           =   0
         Left            =   5505
         TabIndex        =   7
         Tag             =   "1"
         Top             =   420
         Value           =   1  'Checked
         Width           =   2865
      End
      Begin VB.CheckBox OutputStyleOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop &before next action"
         Height          =   780
         Index           =   1
         Left            =   5550
         TabIndex        =   6
         Tag             =   "1"
         Top             =   945
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin VB.ListBox AnimationListBox 
         Height          =   4155
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Tag             =   "1"
         Top             =   390
         Width           =   5280
      End
      Begin ManoharButton.MyButton Command1 
         Height          =   390
         Index           =   0
         Left            =   6105
         TabIndex        =   21
         Tag             =   "1"
         Top             =   3015
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "&Play"
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
         MICON           =   "frmAgentViewer.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton Command1 
         Height          =   390
         Index           =   1
         Left            =   6075
         TabIndex        =   22
         Tag             =   "1"
         Top             =   3765
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "&Stop"
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
         MICON           =   "frmAgentViewer.frx":0060
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
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5655
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   291
   End
   Begin ManoharButton.MyButton Open 
      Height          =   480
      Left            =   300
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1155
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Open Animation File"
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
      MICON           =   "frmAgentViewer.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Advanced 
      Height          =   480
      Left            =   300
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1725
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Advanced Options"
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
      MICON           =   "frmAgentViewer.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton cmdSave 
      Height          =   480
      Left            =   300
      TabIndex        =   28
      Tag             =   "1"
      Top             =   5040
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Save Animations To File"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgentViewer.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton cmdAbout 
      Height          =   480
      Left            =   300
      TabIndex        =   29
      Tag             =   "1"
      Top             =   5610
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "A&bout "
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
      MICON           =   "frmAgentViewer.frx":00D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   4995
      Top             =   105
      _cx             =   847
      _cy             =   847
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
      Caption         =   "Manohar Animation Character Viewer 1.0"
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
      Left            =   945
      TabIndex        =   0
      Tag             =   "1"
      Top             =   75
      Width           =   4095
   End
End
Attribute VB_Name = "frmAgentViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''****************************************************
''   CODED BY: MANOHAR SINGH NEGI                    *
''             6th Semester , I.S.E.                 *
''             R.V. College Of Engineering           *
''             Bangalore - 560059                    *
''             manohar.negi@gmail.com                *
''                                                   *
''****************************************************
'

Dim cc As IAgentCtlCharacterEx
Dim cL As Boolean
Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8

Private Sub Advanced_Click()
MyAgent.PropertySheet.Visible = True
End Sub

Private Sub BalloonStyleOption_Click(Index As Integer)
'-----------------------------------------------------
'-- When the user changes a balloon style option
'-- update the character's word balloon settings
'-----------------------------------------------------

Select Case Index

Case 0 '-- The balloon display option
    If BalloonStyleOption(0).Value = 0 Then
        cc.Balloon.Style = cc.Balloon.Style And (Not BalloonOn)
        BalloonStyleOption(1).Enabled = False
        BalloonStyleOption(2).Enabled = False
        BalloonStyleOption(3).Enabled = False
    Else
        cc.Balloon.Style = cc.Balloon.Style Or BalloonOn
        BalloonStyleOption(1).Enabled = True
        BalloonStyleOption(2).Enabled = True
        BalloonStyleOption(3).Enabled = True
    End If
    
Case 1 '-- The Auto-Hide option
    If BalloonStyleOption(1).Value = 0 Then
        cc.Balloon.Style = cc.Balloon.Style And (Not AutoHide)
    Else
        cc.Balloon.Style = cc.Balloon.Style Or AutoHide
    End If
    
Case 2 '-- The Auto-Pace option
    If BalloonStyleOption(2).Value = 0 Then
        cc.Balloon.Style = cc.Balloon.Style And (Not AutoPace)
    Else
        cc.Balloon.Style = cc.Balloon.Style Or AutoPace
    End If
    
Case 3 '-- The Size-To-Text option
    If BalloonStyleOption(3).Value = 0 Then
        cc.Balloon.Style = cc.Balloon.Style And (Not SizeToText)
    Else
        cc.Balloon.Style = cc.Balloon.Style Or SizeToText
    End If
    
End Select

End Sub


Private Sub cmdAbout_Click()
frmSplashAV.Show vbModal
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If cL = True Then

cc.Hide

With cd2
    .CancelError = True
    .InitDir = App.Path & "\Chars"
    .DialogTitle = "Enter a file name"
    .Filter = "Microsoft Agent Characters Description(*.txt)|*.txt"
    .DefaultExt = "txt"
    .FilterIndex = 1
    .Filename = StrConv(cc.Name, vbProperCase)
    .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
    'Show the Open dialog
    .ShowSave
End With

If (Err = 32755) Then 'operation cancelled
    Beep
    msg = "Operatoin was cancelled" & vbCrLf
    iIndx = MsgBox(msg, vbOKOnly + vbInformation, "BackUp Table Data")
    BackUp = False
    Exit Sub

End If

cc.Show

If cd2.FileTitle <> "" Then
Open cd2.Filename For Output As #2
End If

Print #2,
Print #2, "--------------------------------------------------------------"
Print #2, "List Of Animations For Character : " & cc.Name
Print #2, "--------------------------------------------------------------"
Print #2,

For i = 0 To AnimationListBox.ListCount - 1
    Print #2, AnimationListBox.List(i)
Next

Print #2,
Print #2, "--------------------------------------------------------------"
Print #2, "This list is generated by Manohar Animation Character Viewer 1.0"
Print #2, "--------------------------------------------------------------"
Print #2,

Close #2


x = MsgBox("Created File " & cd2.Filename & vbCrLf & "Do you want to see that file.", vbInformation + vbYesNo, "Save Complete")
If x = vbYes Then
    frmTextEditor.lblTitle.Caption = "Manohar Text Editor - " & cd2.Filename
    frmTextEditor.rtfFile.Locked = True
    frmTextEditor.rtfFile.LoadFile (cd2.Filename)
    frmTextEditor.Show vbModal

End If

End If

End Sub

Private Sub Form_Resize()
If cL = True Then
    If Me.WindowState = vbMinimized Then
        cc.Hide True
    ElseIf Me.WindowState = vbNormal Then
        cc.Show True
    End If
End If

End Sub

Private Sub MyAgent_Command(ByVal UserInput As Object)
'-----------------------------------------------------
'-- If the user selects the Advanced Character Options
'-- command in the character's pop-up menu
'-- make the window visible
'-----------------------------------------------------

If UserInput.Name = "AdvCharOptions" Then
    MyAgent.PropertySheet.Visible = True
End If

End Sub
Private Sub MyAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal Y As Integer)
'-----------------------------------------------------
'-- If the user drags the character
'-- update the character position fields
'-----------------------------------------------------

txtCX.Text = cc.Left
txtCY.Text = cc.Top

End Sub
Private Sub AnimationFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMX.Text = x
txtMY.Text = Y
End Sub

Private Sub AnimationListBox_DblClick()
If cL = True Then
    cc.StopAll
End If
If AnimationListBox.Text = "Show" Then
    cc.Show
    Call cc.MoveTo(CInt(txtCX.Text), CInt(txtCY.Text))
    Exit Sub
End If
Call cc.Play(AnimationListBox.Text)
End Sub

Private Sub AnimationListBox_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMX.Text = x
txtMY.Text = Y

End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
If cL = True Then
    cc.StopAll
End If
If AnimationListBox.Text = "Show" Then
    cc.Show
    Call cc.MoveTo(CInt(txtCX.Text), CInt(txtCY.Text))
    Exit Sub
End If
If AnimationListBox.Text <> "" Then
    Call cc.Play(AnimationListBox.Text)
End If
Case 1
Call cc.StopAll

Case 2 '-- The Speak button was chosen
    
    '-- Speak the text if there is text
    If Not SpeakText.Text = "" Then
        cc.Speak SpeakText.Text
    End If
    
    SpeakText.SetFocus
    SpeakText.SelStart = 0
    SpeakText.SelLength = Len(SpeakText.Text)

Case 3
Call cc.MoveTo(CInt(txtCX.Text), CInt(txtCY.Text))

End Select
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


Private Sub Form_Load()
On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Manohar Animation Character Viewer 1.0"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
'
''----------------------------------------------------------
''-- When the form loads, set the IgnoreSizeEvent flag
''-- (used to differentiate when the Character Animation
''-- Previewer window is restored), set the CharLoaded flag
''-- (used to track when a character is loaded),
''-- and set the initial state of the status bar.
''----------------------------------------------------------
'IgnoreSizeEvent = True
''
''CharLoaded = False

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "000")
txtMX.Text = x
txtMY.Text = Y
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
cL = False
Unload Me
frmMain.Show
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "001")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
'
'
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

Private Sub Open_Click()
On Error Resume Next
If cL = True Then
cc.StopAll
cc.Hide
End If

With cd1
    .CancelError = True
    .InitDir = App.Path & "\Chars\"
    .DialogTitle = "Select a file to open"
    .Filter = "Microsoft Agent Characters (*.acs)|*.acs|Microsoft Agent Characters (*.acg)|*.acg"
    .DefaultExt = "acs"
    .FilterIndex = 1
    .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
    'Show the Open dialog
    .ShowOpen
End With

If cd1.FileTitle = "" Then
    cmdSave.Enabled = False
Else

    MyAgent.Characters.Load curagent & cd1.FileTitle, cd1.Filename
    
    Set cc = MyAgent.Characters(curagent & cd1.FileTitle)
   
    
    Select Case cc.Name

    Case "Merlin"
        Call cc.MoveTo(45, 179)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    Case "Cami"
        Call cc.MoveTo(49, 171)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name

    Case "Genie"
        Call cc.MoveTo(43, 181)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name

    Case "Countney"
        Call cc.MoveTo(68, 209)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "The Genius"
        Call cc.MoveTo(45, 198)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    
    Case "Earl"
        Call cc.MoveTo(65, 202)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    
    Case "James"
        Call cc.MoveTo(34, 151)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "Max"
        Call cc.MoveTo(47, 179)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "Robby"
        Call cc.MoveTo(49, 179)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "Rover"
        Call cc.MoveTo(71, 205)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "VRGirl"
        Call cc.MoveTo(47, 182)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "The Dot"
        Call cc.MoveTo(46, 179)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
        
    Case "F1"
        Call cc.MoveTo(44, 184)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
                
    Case "Office Logo"
        Call cc.MoveTo(44, 193)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    
    Case "Mother Nature"
        Call cc.MoveTo(41, 196)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
       
    Case "Rocky"
        Call cc.MoveTo(43, 180)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    
    Case "Clippit"
        Call cc.MoveTo(44, 193)
        MsgBox "Loaded Animation Character : " & cc.Name, vbInformation + vbOKOnly, "Loaded " & cc.Name
    
    End Select
'
    cc.Show
    cL = True
    Call cc.Stop
    txtCX.Text = cc.Left
    txtCY.Text = cc.Top
    Call cc.Play("Announce")
    Call LoadAnimation
    InitPopupMenuCmds
    AnimationFrame.Enabled = True
    SpeechOut.Enabled = True
    Frame3.Enabled = True
    Me.Caption = cc.Name + " - Manohar Animation Character Viewer 1.0"
    lblTitle.Caption = Me.Caption
    cmdSave.Enabled = True
    
End If

'*****************************
'--- user cancels operation
'*****************************
If (Err = 32755) Then 'operation cancelled
    Beep
    msg = "Operatoin was cancelled" & vbCrLf
    iIndx = MsgBox(msg, vbOKOnly + vbInformation, "BackUp Table Data")
    BackUp = False
    Exit Sub
Else
'    On Error GoTo expError
End If

End Sub

Sub LoadAnimation()
'-- Load the character's animation into the list box
AnimationListBox.CLEAR
For Each AnimationName In cc.AnimationNames
        AnimationListBox.AddItem AnimationName
Next
End Sub
Sub InitPopupMenuCmds()
'-----------------------------------------------------
'-- Add a command to the character to provide access
'-- to the Advanced Character Options
'-----------------------------------------------------
cc.Commands.RemoveAll
cc.Commands.Add "AdvCharOptions", "&Advanced Character Options"

End Sub

Private Sub OutputStyleOption_Click(Index As Integer)
'-----------------------------------------------------
'-- If the Play Sound Effects option is changed
'-- set the character's output option
'-----------------------------------------------------

If Index = 0 Then
    If OutputStyleOption(0).Value Then
        cc.SoundEffectsOn = True
        
    Else
        cc.SoundEffectsOn = False
        
    End If
   
End If

End Sub

Private Sub SpeakText_GotFocus()
'--------------------------------------------------
'-- If the user clicks on this text box and
'-- it's enabled make The Speak button the
'-- default button
'--------------------------------------------------
If Command1(2).Enabled Then
    Command1(2).Default = True
    SpeakText.SelStart = 0
    SpeakText.SelLength = Len(SpeakText.Text)
End If

End Sub

Private Sub SpeakText_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMX.Text = x
txtMY.Text = Y

End Sub

Private Sub SpeechOut_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtMX.Text = x
txtMY.Text = Y
End Sub
