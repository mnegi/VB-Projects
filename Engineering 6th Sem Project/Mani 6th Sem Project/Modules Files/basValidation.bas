Attribute VB_Name = "basValidation"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Public Function NAMEVALID(code As Integer)
If Chr(code) Like "[A-Z a-z]" Or code = 8 Or code = 32 Or code = 13 Then
Exit Function
Else
code = 0
End If
End Function
Public Function NUMVALID(code As Integer)
If Chr(code) Like "[0-9]" Or code = 8 Or code = 32 Or code = 13 Then
Exit Function
Else
code = 0
End If
End Function
Public Function DATEVALID(code As Integer)
If Chr(code) Like "[0-9]" Or code = 8 Or code = 13 Or Chr(code) Like "[A-Z a-z]" Or Chr(code) Like "-" Then
Exit Function
Else
code = 0
End If
End Function
Public Function CHARVALID(code As Integer)
If Chr(code) Like "[A-Z a-z]" Or code = 35 Or code = 64 Or code = 45 Or code = 46 Or code = 8 Or code = 36 Then
Exit Function
Else
code = 0
End If
End Function
Public Function ADDVALID(code As Integer)
If Chr(code) Like "[A-Z a-z 0-9]" Or code = 95 Or code = 42 Or code = 126 Or code = 8 Or code = 64 Or code = 35 Or code = 36 Or code = 40 Or code = 41 Or code = 45 Or code = 92 Or code = 123 Or code = 91 Or code = 125 Or code = 93 Or code = 44 Or code = 46 Or code = 47 Or code = 13 Then
Exit Function
Else
code = 0
End If
End Function
Public Function EMAILVALID(code As Integer)
If Chr(code) Like "[A-Z a-z 0-9]" Or code = 8 Or code = 95 Or code = 42 Or code = 126 Or code = 64 Or code = 35 Or code = 36 Or code = 40 Or code = 41 Or code = 45 Or code = 92 Or code = 123 Or code = 91 Or code = 125 Or code = 93 Or code = 44 Or code = 46 Or code = 47 Or code = 33 Or code = 37 Or code = 94 Or code = 38 Or code = 124 Or code = 62 Or code = 60 Or code = 63 Then
Exit Function
Else
code = 0
End If
End Function
Public Function PHVALID(code As Integer)
If Chr(code) Like "[0-9]" Or code = 45 Or code = 8 Then
Exit Function
Else
code = 0
End If
End Function

Public Function ALNUMVALID(code As Integer)
If Chr(code) Like "[A-Z a-z 0-9]" Or code = 45 Or code = 8 Then
Exit Function
Else
code = 0
End If
End Function

