Attribute VB_Name = "ErrorHandling"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Public Sub Handle_Error(title As String, err_no As String, err_desc As String, err_ico As String, appIcon As String, err_but As Integer, TextBox As Integer)
    frmMsgbox.lblTitle.Caption = title
    frmMsgbox.lblError.Caption = err_no
    frmMsgbox.lblDesc.Caption = err_desc
    frmMsgbox.imgAppIcon.Picture = LoadPicture(App.Path & "\Common\Icons\App Icons\" & appIcon)
    frmMsgbox.imgIcon.Picture = LoadPicture(App.Path & "\Common\Images\" & err_ico)
    Select Case err_but
    Case 1
        frmMsgbox.cmdOkOnly.Visible = True
    Case 2
        frmMsgbox.cmdCancel.Visible = True
        frmMsgbox.cmdok.Visible = True
    End Select
   If TextBox = 0 Then
   frmMsgbox.txtMsgbox.Visible = False
   frmMsgbox.lblDesc.Visible = True
   Else
   frmMsgbox.txtMsgbox.Visible = True
   frmMsgbox.lblDesc.Visible = False
   End If
   
End Sub
    


