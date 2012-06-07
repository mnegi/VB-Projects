Attribute VB_Name = "basExport"

Public Function CSV(adorecordset As ADODB.Recordset, Filename As String, FolderName As String, UseFolder As Boolean) As Boolean
Dim iTotalRecords As Integer
Dim sFileToExport As String
Dim iFileNum As Integer
Dim msg As String
Dim iIndx As Integer
Dim iNumnberOfFields As Integer
If UseFolder = False Then

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

