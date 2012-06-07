VERSION 5.00
Begin VB.Form frmImagePreview 
   BackColor       =   &H00F8F8F8&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmImagePreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgPreview 
      Height          =   8640
      Left            =   0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   12000
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   11685
      Stretch         =   -1  'True
      Top             =   30
      Width           =   285
   End
End
Attribute VB_Name = "frmImagePreview"
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

Private Sub Form_Load()
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\c_up.jpg")
If MyFile.FileExists(ImagePath) Then
imgPreview.Picture = LoadPicture(ImagePath)
End If
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\c_up.jpg")
End Sub

Private Sub Form_Terminate()
ImagePath = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
ImagePath = ""
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\c_down.jpg")
End Sub
