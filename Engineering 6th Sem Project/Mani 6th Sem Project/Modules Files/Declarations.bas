Attribute VB_Name = "Declarations"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Const SW_SHOWMAXIMIZED = 3
'variable declaration for themes
'theme name
Public theme As String
'label settings
Public LabelForecolor As String
Public LabelFontname As String
Public LabelFontsize As String
Public LabelFontbold As String
'textbox & combobox settings
Public TextForecolor As String
Public TextBackcolor As String
Public TextLinecolor As String
Public TextFontname As String
Public TextFontsize As String
Public TextFontbold As String
'Mybuttons setting i.e. for command buttons as simple
Public BtnForecolor As String
Public BtnForeover As String
Public BtnBackcolor As String
Public BtnBackover As String
Public BtnColorscheme As String
Public BtnFontname As String
Public BtnFontsize As String
Public BtnFontbold As String
Public BtnColorschemes As String
'BtnColorschemes can have values 0 to 3

Public MenuLabelFontSize As String
Public MenuLabelBackColor As String
Public MenuLabelFontName As String
Public MenuFrameColor As String

'if textbox used as password box
Public TextPasswordFont As String
Public TextPasswordChar As String

'image to prewiew
Public ImagePath As String


'open filesystem object

Public MyFile As New FileSystemObject


Public username As String

Public Cn As New ADODB.Connection
Public Cmd As New ADODB.Command
Public rs As New ADODB.Recordset
Public Par As New ADODB.Parameter

'msgbox return true if ok is clicked false otherwise
Public MsgBOx_R_Value As Boolean
Public MsgBox_R_Text As String

Public cc As IAgentCtlCharacter
Public characterlocation As String

Public RegPath As String

Sub main()
On Error Resume Next
username = GetSetting(App.title, "Users", "Name", "None")
theme = GetSetting(App.title, "Theme", "Name", "Gray")
If username = "None" Then
frmUserCreation.Show
Else
Call Get_Theme
frmLogin.Show
End If
End Sub

Public Function SetWallpaper(Filename As String) As Boolean
Dim rc As Long
'call the api
rc = SystemParametersInfo(SPI_SETDESKWALLPEPER, O&, ByVal Filename, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

'RETURN RESULT
SetWallpaper = (rc = 0)
End Function

