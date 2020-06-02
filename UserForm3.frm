VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub RElementF_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Fire", _
        Prompt:="Select Area to Remove Fire:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    If Cell.Interior.Color = RGB(255, 200, 0) And _
        Cell.Interior.Pattern = xlPatternChecker And _
        Cell.Interior.PatternColor = vbRed _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        Cell.Borders.LineStyle = xlDot
    Else
    End If

Next Cell

UserForm3.RElementF.Value = False

End Sub

Private Sub RElementR_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Rocks", _
        Prompt:="Select Area to Remove Rocks:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    If Cell.Interior.Color = RGB(166, 166, 166) And _
        Cell.Interior.Pattern = xlPatternGrid And _
        Cell.Interior.PatternColor = vbBlack _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        Cell.Borders.LineStyle = xlDot
    End If
Next Cell

UserForm3.RElementR.Value = False

End Sub

Private Sub RElementS_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Sand", _
        Prompt:="Select Area to Remove Sand:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    If Cell.Interior.Color = RGB(255, 255, 183) And _
        Cell.Interior.Pattern = xlPatternGray16 And _
        Cell.Interior.PatternColor = RGB(204, 153, 0) _
    Then
        Cell.Interior.Pattern = none
    End If
Next Cell

UserForm3.RElementR.Value = False


End Sub

Private Sub RElementT_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Trees", _
        Prompt:="Select Area to Remove Trees:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    If Cell.Interior.Color = RGB(84, 130, 53) And _
        Cell.Interior.Pattern = xlSolid _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        Cell.Borders.LineStyle = xlDot
    End If
    
Next Cell

UserForm3.RElementT.Value = False

End Sub

Private Sub RElementW_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Water", _
        Prompt:="Select Area to Remove Water:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    If Cell.Interior.Color = RGB(0, 176, 240) And _
        Cell.Interior.Pattern = xlPatternGray16 And _
        Cell.Interior.PatternColor = vbBlue _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        Cell.Borders.LineStyle = xlDot
    End If
Next Cell

UserForm3.RElementW.Value = False

End Sub

Private Sub RElementWood_Click()

UserForm3.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Wood", _
        Prompt:="Select Area to Remove Wood:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral

    If Cell.Interior.Color = RGB(128, 96, 0) And _
        Cell.Interior.Pattern = xlPatternLightDown And _
        Cell.Interior.PatternColor = vbBlack _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        Cell.Borders.LineStyle = xlDot
        
    End If
Next Cell

UserForm3.RElementWood.Value = False

End Sub
