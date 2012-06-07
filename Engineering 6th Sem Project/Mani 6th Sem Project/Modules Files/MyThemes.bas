Attribute VB_Name = "MyThemes"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Public Sub Get_Theme()

theme = GetSetting(App.title, "Theme", "Name", "Gray")
LabelForecolor = GetSetting(App.title, "Theme", "LabelForecolor")
LabelFontname = GetSetting(App.title, "Theme", "LabelFontname")
LabelFontsize = GetSetting(App.title, "Theme", "LabelFontsize")
LabelFontbold = GetSetting(App.title, "Theme", "LabelFontbold")


'textbox & combobox settings
TextBackcolor = GetSetting(App.title, "Theme", "TextBackcolor")
TextForecolor = GetSetting(App.title, "Theme", "TextForecolor")
TextLinecolor = GetSetting(App.title, "Theme", "TextLinecolor")
TextFontname = GetSetting(App.title, "Theme", "TextFontname")
TextFontsize = GetSetting(App.title, "Theme", "TextFontsize")
TextFontbold = GetSetting(App.title, "Theme", "TextFontbold")
TextPasswordFont = GetSetting(App.title, "Theme", "TextPasswordFont")
TextPasswordChar = GetSetting(App.title, "Theme", "TextPasswordChar")

'Mybuttons setting i.e. for command buttons as simple
BtnColorscheme = GetSetting(App.title, "Theme", "BtnColorscheme")
BtnForecolor = GetSetting(App.title, "Theme", "BtnForecolor")
BtnForeover = GetSetting(App.title, "Theme", "BtnForeover")
BtnBackcolor = GetSetting(App.title, "Theme", "BtnBackcolor")
BtnBackover = GetSetting(App.title, "Theme", "BtnBackover")

BtnFontname = GetSetting(App.title, "Theme", "BtnFontname")
BtnFontsize = GetSetting(App.title, "Theme", "BtnFontsize")
BtnFontbold = GetSetting(App.title, "Theme", "BtnFontbold")

' Menu settings
MenuLabelFontSize = GetSetting(App.title, "Theme", "MenuLabelFontSize")
MenuLabelFontName = GetSetting(App.title, "Theme", "MenuLabelFontName")
MenuLabelBackColor = GetSetting(App.title, "Theme", "MenuLabelBackColor")
MenuFrameColor = GetSetting(App.title, "Theme", "MenuFrameColor")

End Sub
Public Sub Apply_Theme(fn As Form, FormSize As Integer)

'fn.MousePointer = 99
'fn.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Select Case FormSize
Case 0
'add the small image
fn.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\small1.jpg")

Case 1
'add the userScreen image
fn.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\1.jpg")

Case 2
'add the add photo size image
fn.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\window\addimage.jpg")

Case 3
'add the maximised image
fn.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\window\maximised.jpg")

Case 4
'add the restore image
fn.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\window\restore.jpg")

End Select
'call the buttons funtion to show the close,min,restore
Call ShowButtons(fn, "000")


With fn

For iindex = 0 To .Controls.Count - 1
If .Controls(iindex).Tag = "1" Then

If (TypeOf .Controls(iindex) Is Text) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).LineColor = TextLinecolor
    .Controls(iindex).FontSize = TextFontsize
    .Controls(iindex).FontName = TextFontname
    .Controls(iindex).FontBold = TextFontbold
End If

If (TypeOf .Controls(iindex) Is TextBox) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).FontSize = TextFontsize
    .Controls(iindex).FontName = TextFontname
    .Controls(iindex).FontBold = TextFontbold
End If

If (TypeOf .Controls(iindex) Is RichTextBox) Then
    .Controls(iindex).BackColor = TextBackcolor
'    .Controls(iindex).FontBold = TextFontbold
End If

If (TypeOf .Controls(iindex) Is ComboBox Or TypeOf .Controls(iindex) Is ListBox) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).FontName = TextFontname
    .Controls(iindex).FontSize = TextFontsize
    .Controls(iindex).FontBold = TextFontbold
    .Controls(iindex).MousePointer = 99
    .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Pen.cur")
End If
If (TypeOf .Controls(iindex) Is ImageCombo) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).MousePointer = 99
    .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Pen.cur")
End If

If (TypeOf .Controls(iindex) Is Label) Then
    .Controls(iindex).ForeColor = LabelForecolor
    .Controls(iindex).FontName = LabelFontname
    .Controls(iindex).FontBold = LabelFontbold
End If
If (TypeOf .Controls(iindex) Is MyButton) Then
    .Controls(iindex).ForeColor = BtnForecolor
    .Controls(iindex).ForeOver = BtnForeover
    .Controls(iindex).BackColor = BtnBackcolor
    .Controls(iindex).BackOver = BtnBackover
    .Controls(iindex).FontName = BtnFontname
    .Controls(iindex).FontSize = BtnFontsize
    .Controls(iindex).FontBold = BtnFontbold
    .Controls(iindex).ColorScheme = BtnColorscheme
    .Controls(iindex).MousePointer = 99
    .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Link.cur")


End If
If (TypeOf .Controls(iindex) Is CheckBox) Then
    .Controls(iindex).ForeColor = LabelForecolor
    .Controls(iindex).MaskColor = TextBackcolor
    .Controls(iindex).FontName = TextFontname
    .Controls(iindex).FontSize = TextFontsize
    .Controls(iindex).FontBold = TextFontbold
    .Controls(iindex).MousePointer = 99
    .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Pen.cur")

End If
If (TypeOf .Controls(iindex) Is OptionButton) Then
    .Controls(iindex).ForeColor = LabelForecolor
    .Controls(iindex).FontName = LabelFontname
    .Controls(iindex).BackColor = TextBackcolor
'    .Controls(iindex).FontSize = LabelFontsize
'    .Controls(iindex).FontBold = LabelFontbold
    .Controls(iindex).MousePointer = 99
    .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Pen.cur")

End If

'
If (TypeOf .Controls(iindex) Is MaskEdBox) Then
        .Controls(iindex).BackColor = TextBackcolor
        .Controls(iindex).ForeColor = TextForecolor
        .Controls(iindex).FontName = TextFontname
        .Controls(iindex).FontSize = TextFontsize
        .Controls(iindex).FontBold = TextFontbold
        .Controls(iindex).MousePointer = 99
        .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Beam.cur")
End If
'
If (TypeOf .Controls(iindex) Is Frame) Then
        .Controls(iindex).ForeColor = LabelForecolor
        .Controls(iindex).FontName = LabelFontname
        .Controls(iindex).FontBold = LabelFontbold
'
End If
If (TypeOf .Controls(iindex) Is TreeView) Then
        .Controls(iindex).MousePointer = 99
        .Controls(iindex).MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\pen.cur")
End If


'
'
End If

If .Controls(iindex).Tag = "2" Then

If (TypeOf .Controls(iindex) Is Text) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).LineColor = TextLinecolor
    .Controls(iindex).FontSize = 10
    
End If

If (TypeOf .Controls(iindex) Is TextBox) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).FontSize = 10
    
End If

If (TypeOf .Controls(iindex) Is ComboBox) Then
    .Controls(iindex).BackColor = TextBackcolor
    .Controls(iindex).ForeColor = TextForecolor
    .Controls(iindex).Font = TextFontname
    .Controls(iindex).FontSize = 10
    
End If

If (TypeOf .Controls(iindex) Is Label) Then
    .Controls(iindex).BackColor = MenuLabelBackColor
    .Controls(iindex).FontName = MenuLabelFontName
    .Controls(iindex).FontSize = MenuLabelFontSize
End If

If (TypeOf .Controls(iindex) Is Frame) Then
    .Controls(iindex).BackColor = MenuFrameColor
End If

If (TypeOf .Controls(iindex) Is RichTextBox) Then
    .Controls(iindex).BackColor = TextBackcolor
'    .Controls(iindex).FontName = "Baraha Devanagari Extra"
'    .Controls(iindex).FontBold = TextFontbold
End If

End If
Next
End With


End Sub

Public Sub Change_Theme(theme_name As String)
theme = theme_name
 
Select Case theme
Case "Gray"
    
            
    LabelForecolor = 4210752
    LabelFontname = "Times New Roman"
    LabelFontsize = 12
    LabelFontbold = "True"
    
    TextBackcolor = &HEAEAEA
    TextForecolor = &H400000
    TextLinecolor = &H0
    TextFontname = "Palatino Linotype"
    TextFontsize = 12
    TextFontbold = "True"
    TextPasswordFont = "Wingdings"
    TextPasswordChar = "["
    
    
    BtnColorscheme = 2
    BtnForecolor = &H400000
    BtnForeover = &H404040
    BtnBackcolor = &HC0C0C0
    BtnBackover = &HE0E0E0
    BtnFontname = "Palatino Linotype"
    BtnFontsize = 12
    BtnFontbold = "True"
    
    MenuLabelFontSize = 11
    MenuFrameColor = &H8000000F
    MenuLabelBackColor = &HC0C0C0
    MenuLabelFontName = "Bell MT"
    

Case "Blue"
    
    LabelForecolor = &H400000
    LabelFontname = "Palatino Linotype"
    LabelFontsize = 12
    LabelFontbold = "True"
    
    TextBackcolor = &HFDFEF3
    TextForecolor = &H0
    TextLinecolor = &HB1AA54
    TextFontname = "Palatino Linotype"
    TextFontsize = 12
    TextFontbold = "True"
    TextPasswordFont = "Wingdings"
    TextPasswordChar = "]"
    
    BtnColorscheme = 2
    BtnForecolor = &H404000
    BtnForeover = &H400000
    BtnBackcolor = &HFFFFC0
    BtnBackover = &HC0E0FF
    BtnFontname = "MS Serif"
    BtnFontsize = 10
    BtnFontbold = "True"

    MenuLabelFontSize = 11
    MenuFrameColor = &HFEFDD6
    MenuLabelBackColor = &HFDE7B3
    MenuLabelFontName = "Bell MT"
    
Case "Red"

    LabelForecolor = &H400000
    LabelFontname = "Times New Roman"
    LabelFontsize = 12
    LabelFontbold = "True"
    
    TextBackcolor = &HF0F0FF
    TextForecolor = &HC0&
    TextLinecolor = &H40C0
    TextFontname = "Times New Roman"
    TextFontsize = 12
    TextFontbold = "True"
    TextPasswordFont = "Webdings"
    TextPasswordChar = "Y"
       
    BtnColorscheme = 2
    BtnForecolor = &H80
    BtnForeover = &H40C0
    BtnBackcolor = &HC0C0FF
    BtnBackover = &HFFC0FF
    BtnFontname = "Georgia"
    BtnFontsize = 10
    BtnFontbold = "True"
    
    MenuLabelFontSize = 11
    MenuFrameColor = &HC0C0FF
    MenuLabelBackColor = &HE2E3FE
    MenuLabelFontName = "Bell MT"
    
    
Case "Green"
    LabelForecolor = &H404000
    LabelFontname = "Georgia"
    LabelFontsize = 10
    LabelFontbold = "True"
    
    TextBackcolor = &HDFFFE1
    TextForecolor = &H404000
    TextLinecolor = &H80FF80
    TextFontname = "Palatino Linotype"
    TextFontsize = 12
    TextFontbold = "True"
    TextPasswordFont = "Wingdings"
    TextPasswordChar = "T"
    
    BtnColorscheme = 2
    BtnForecolor = &H400000
    BtnForeover = &H808000
    BtnBackcolor = &HC0FFC0
    BtnBackover = &HCEF8FF
    BtnFontname = "Palatino Linotype"
    BtnFontsize = 12
    BtnFontbold = "True"
    
    MenuLabelFontSize = 11
    MenuFrameColor = &HD0EDC5
    MenuLabelBackColor = &HDEFAEB
    MenuLabelFontName = "Bell MT"
    
End Select
End Sub
Public Sub Save_Theme()
   SaveSetting App.title, "Theme", "LabelForecolor", LabelForecolor
   SaveSetting App.title, "Theme", "LabelFontname", LabelFontname
   SaveSetting App.title, "Theme", "LabelFontbold", LabelFontbold
   
   SaveSetting App.title, "Theme", "TextBackcolor", TextBackcolor
   SaveSetting App.title, "Theme", "TextForecolor", TextForecolor
   SaveSetting App.title, "Theme", "TextLinecolor", TextLinecolor
   SaveSetting App.title, "Theme", "TextFontname", TextFontname
   SaveSetting App.title, "Theme", "TextFontsize", TextFontsize
   SaveSetting App.title, "Theme", "TextFontbold", TextFontbold
   SaveSetting App.title, "Theme", "TextPasswordChar", TextPasswordChar
   SaveSetting App.title, "Theme", "TextPasswordFont", TextPasswordFont
    
   SaveSetting App.title, "Theme", "BtnColorscheme", BtnColorscheme
   SaveSetting App.title, "Theme", "BtnForecolor", BtnForecolor
   SaveSetting App.title, "Theme", "BtnForeover", BtnForeover
   SaveSetting App.title, "Theme", "BtnBackcolor", BtnBackcolor
   SaveSetting App.title, "Theme", "BtnBackover", BtnBackover
   SaveSetting App.title, "Theme", "BtnFontname", BtnFontname
   SaveSetting App.title, "Theme", "BtnFontsize", BtnFontsize
   SaveSetting App.title, "Theme", "BtnFontbold", BtnFontbold
   
   SaveSetting App.title, "Theme", "MenuLabelFontName", MenuLabelFontName
   SaveSetting App.title, "Theme", "MenuLabelFontSize", MenuLabelFontSize
   SaveSetting App.title, "Theme", "MenuLabelBackColor", MenuLabelBackColor
   SaveSetting App.title, "Theme", "MenuFrameColor", MenuFrameColor
End Sub

Sub ShowButtons(fn As Form, btn As String)

If Mid$(btn, 1, 1) = "0" Then
fn.imgMin.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\M_up.jpg")
Else
fn.imgMin.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\M_down.jpg")
End If

If Mid$(btn, 2, 1) = "0" Then
fn.imgRestore.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\R_up.jpg")
Else
fn.imgRestore.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\R_down.jpg")
End If

If Mid$(btn, 3, 1) = "0" Then
fn.imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\C_up.jpg")
Else
fn.imgClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Window\C_down.jpg")
End If
End Sub


Sub Load_PasswordChar(TextBoxName As Text)
TextBoxName.FontSize = 12
TextBoxName.FontName = TextPasswordFont
TextBoxName.PasswordChar = TextPasswordChar
End Sub
