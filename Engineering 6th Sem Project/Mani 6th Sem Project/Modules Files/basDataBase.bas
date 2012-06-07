Attribute VB_Name = "basDataBase"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************



Public Type FieldTypes
    ObjectType As String * 10
    ObjectName As String * 50
    FieldName As String * 50
    FieldDataType As String * 50
    FieldNull As String * 10
    FieldConstraint As String * 100
End Type



Public Function CreateTable(x As String) As String

'create the database table
Cmd.CommandText = x
On Error GoTo EXIT_SUB

Cmd.Execute

'save the command or structure of the command to file

'get the name of the table

Dim Scount, tnStart, tnEnd As Integer
Scount = 0

For i = 1 To Len(x) - 1
If Mid(x, i, 1) = " " Or Mid(x, i, 1) = "(" Then
    Scount = Scount + 1
    If Scount = 2 Then
        tnStart = i + 1
    End If
    If Scount = 3 Then
        tnEnd = i
    End If
End If
Next

'now we got the table name as mid(X,tnStart,tnEnd-tnstart)
'open the file for that table

Open App.Path & "\Tables\" & Mid(x, tnStart, tnEnd - tnStart) & ".txt" For Output As #2

'print the syntax like "create table <table>" to the file

Print #2, Mid(x, 1, tnEnd - 1)
Print #2, " "
Print #2, "   ("
Print #2, " "

'now save the fields to the file

Dim LastCount, Count
LastCount = tnEnd + 1
Count = 0

For i = tnEnd + 1 To Len(x) - 1
Count = Count + 1
If Mid(x, i, 1) = "," Or i = Len(x) - 1 Then
    Print #2, "      " & Mid(x, LastCount, Count)
    Count = 0
    LastCount = i + 1
End If
Next

Print #2, " "
Print #2, "   )"
Print #2, " "

Close #2
CreateTable = "1" & Mid(x, tnStart, tnEnd - tnStart)

EXIT_SUB:

If Err.Number <> 0 Then
CreateTable = "0ERROR : " & Err.Number & " : " & Err.Description
End If

End Function

Public Function ReturnTableName(x As String) As String

'get the name of the table

Dim Scount, tnStart, tnEnd As Integer
Scount = 0

For i = 1 To Len(x)
If Mid(x, i, 1) = " " Or Mid(x, i, 1) = "(" Then
    Scount = Scount + 1
    If Scount = 2 Then
        tnStart = i + 1
    End If
    If Scount = 3 Then
        tnEnd = i
        ReturnTableName = Mid(x, tnStart, tnEnd - tnStart)
    End If
End If
Next


End Function
Public Function ReturnTableNameEx(x As String, Num As Integer) As String

'get the name of the table

Dim Scount, tnStart, tnEnd As Integer
Scount = 0

For i = 1 To Len(x)
If Mid(x, i, 1) = " " Or Mid(x, i, 1) = "(" Then
    Scount = Scount + 1
    If Scount = Num - 1 Then
        tnStart = i + 1
    End If
    If Scount = Num Then
        tnEnd = i
        ReturnTableNameEx = Mid(x, tnStart, tnEnd - tnStart)
    End If
End If
Next


End Function

Public Function ReturnCommandName(x As String) As String

'get the name of the table

Dim Count, tEnd As Integer
Count = 0

For i = 1 To Len(x)
If Mid(x, i, 1) = " " Then
    Scount = Scount + 1
    If Scount = 1 Then
        tEnd = i - 1
        ReturnCommandName = Mid(x, 1, tEnd)
    End If
End If
Next


End Function


Sub DeleteAllTables()


On Error GoTo handle
Cmd.ActiveConnection = Cn

'tables of insurance database
Cmd.CommandText = "delete from participated"
Cmd.Execute
Cmd.CommandText = "delete from owns"
Cmd.Execute
Cmd.CommandText = "delete from accident"
Cmd.Execute
Cmd.CommandText = "delete from cars"
Cmd.Execute
Cmd.CommandText = "delete from carmodels"
Cmd.Execute
Cmd.CommandText = "delete from carmakes"
Cmd.Execute
Cmd.CommandText = "delete from manufacturer"
Cmd.Execute
Cmd.CommandText = "delete from person"
Cmd.Execute

'tables of order processing database
Cmd.CommandText = "delete from shipment"
Cmd.Execute
Cmd.CommandText = "delete from warehouse"
Cmd.Execute
Cmd.CommandText = "delete from order_item"
Cmd.Execute
Cmd.CommandText = "delete from order_tab"
Cmd.Execute
Cmd.CommandText = "delete from item"
Cmd.Execute
Cmd.CommandText = "delete from cust"
Cmd.Execute

'tables of student enrollment database
Cmd.CommandText = "delete from book_adoption"
Cmd.Execute
Cmd.CommandText = "delete from enroll"
Cmd.Execute
Cmd.CommandText = "delete from text"
Cmd.Execute
Cmd.CommandText = "delete from course"
Cmd.Execute
Cmd.CommandText = "delete from student"
Cmd.Execute

'tables of book dealer database
Cmd.CommandText = "delete from order_details"
Cmd.Execute
Cmd.CommandText = "delete from catalog"
Cmd.Execute
Cmd.CommandText = "delete from category"
Cmd.Execute
Cmd.CommandText = "delete from publisher"
Cmd.Execute
Cmd.CommandText = "delete from author"
Cmd.Execute

'tables of bank enterprise database
Cmd.CommandText = "delete from borrower"
Cmd.Execute
Cmd.CommandText = "delete from loan"
Cmd.Execute
Cmd.CommandText = "delete from depositor"
Cmd.Execute
Cmd.CommandText = "delete from customer"
Cmd.Execute
Cmd.CommandText = "delete from account"
Cmd.Execute
Cmd.CommandText = "delete from branch"
Cmd.Execute

MsgBox "All the tables are emptied"

handle:

If Err.Number <> 0 Then
MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
Resume Next
End If

End Sub


