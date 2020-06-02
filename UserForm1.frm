VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3420
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AElementF_Click()

UserForm1.Hide

Dim FireArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Fire", _
        Prompt:="Select Fire Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(255, 200, 0)
    Cell.Interior.Pattern = xlPatternChecker
    Cell.Interior.PatternColor = vbRed
    Cell.Borders.LineStyle = none
    
Next Cell

UserForm1.AElementF.Value = False

End Sub

Private Sub AElementR_Click()

UserForm1.Hide

Dim RockArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Rocks", _
        Prompt:="Select Rock Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(166, 166, 166)
    Cell.Interior.Pattern = xlPatternGrid
    Cell.Interior.PatternColor = vbBlack
    Cell.Borders.LineStyle = none
Next Cell

UserForm1.AElementR.Value = False

End Sub

Private Sub AElementS_Click()

UserForm1.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Sand", _
        Prompt:="Select Sand Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(255, 255, 183)
    Cell.Interior.Pattern = xlPatternGray16
    Cell.Interior.PatternColor = RGB(204, 153, 0)
    Cell.Borders.LineStyle = xlDot

Next Cell

UserForm1.AElementS.Value = False


End Sub

Private Sub AElementT_Click()

UserForm1.Hide

Dim TreesArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Trees", _
        Prompt:="Select Tree Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(84, 130, 53)
    Cell.Interior.Pattern = xlSolid
    Cell.Borders.LineStyle = none
    
Next Cell

UserForm1.AElementT.Value = False

End Sub

Private Sub AElementW_Click()

UserForm1.Hide

Dim WaterArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Water", _
        Prompt:="Select Water Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(0, 176, 240)
    Cell.Interior.Pattern = xlPatternGray16
    Cell.Interior.PatternColor = vbBlue
    Cell.Borders.LineStyle = none
    
Next Cell

UserForm1.AElementW.Value = False

End Sub

Private Sub AElementWood_Click()

UserForm1.Hide

Dim WaterArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Adding Wood", _
        Prompt:="Select Wood Positions:", _
        Type:=8)
        
Dim Cell As Range

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(128, 96, 0)
    Cell.Interior.Pattern = xlPatternLightDown
    Cell.Interior.PatternColor = vbBlack
    Cell.Borders.LineStyle = none
    
Next Cell

UserForm1.AElementWood.Value = False

End Sub
