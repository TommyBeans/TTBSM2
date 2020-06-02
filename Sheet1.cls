VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim Numeral2 As Integer

Private Sub AEnemy_Click()

Dim POS() As Variant
Dim Numeral As Variant
Dim Playernames() As Variant

Dim Psheet As Worksheet
Dim PLast As Long

Set Psheet = ThisWorkbook.Worksheets("PlayerSheet")

Numeral = InputBox("Enter Number of Enemies:", "Adding Baddies")

ReDim POS(Numeral - 1)
ReDim Playernames(Numeral - 1)
Dim TopVar As Integer

Dim i As Integer

Application.ScreenUpdating = False

For i = 0 To Numeral - 1

    Set POS(i) = Application.InputBox( _
        Title:="Adding Enemies", _
        Prompt:="Select Enemy(" & i + 1 & ") Position:", _
        Type:=8)
        
    POS(i).Interior.Color = RGB(200, 0, 0)
    
    POS(i).HorizontalAlignment = xlCenterAcrossSelection
    POS(i).Interior.Pattern = xlSolid
        
    POS(i).Value = "ENE (" & i + 1 & ")" & "    "
        
    Psheet.Cells(i + 1, 3) = POS(i)

Next i


TopVar = 10
    
'For i = 0 To Numeral - 1
'    Set Optbtn(i) = UserForm3.Controls.Add("Forms.OptionButton.1")
'    Optbtn(i).Caption = POS(i).Value
'    Optbtn(i).Left = 15
'    Optbtn(i).Top = TopVar
'
'    TopVar = TopVar + 15
'
'Next i

End Sub

Private Sub AFriend_Click()

On Error GoTo 0

Dim Numeral As Variant
Dim Rcolor As Long
Dim Bcolor As Long
Dim Gcolor As Long
Dim NameF As Long

Dim Optionlist() As Variant
Dim i As Integer

Dim Psheet As Worksheet
Dim PLast As Long

Set Psheet = ThisWorkbook.Worksheets("PlayerSheet")

Dim POS() As Variant
Dim Optbtn() As Variant

Numeral = InputBox("Enter Number of Players (+Friendly NPC's):", "Adding Crew")

Numeral2 = Numeral

Dim Playernames() As Variant

ReDim Playernames(Numeral - 1)
ReDim POS(Numeral - 1)
ReDim Optionlist(Numeral - 1)
ReDim Optbtn(Numeral - 1)

Application.ScreenUpdating = False

For i = 0 To Numeral - 1

    Rcolor = WorksheetFunction.RandBetween(175, 255)
    Bcolor = WorksheetFunction.RandBetween(175, 255)
    Gcolor = WorksheetFunction.RandBetween(175, 255)

    Set POS(i) = Application.InputBox( _
        Title:="Adding Crew", _
        Prompt:="Select Crew(" & i + 1 & ") position:", _
        Type:=8)
    
    POS(i).Interior.Color = RGB(Rcolor, Bcolor, Gcolor)
    
    Playernames(i) = InputBox("Name Crew Member:", "Adding Crew") & "    "
    
    Optionlist(i) = Playernames(i)
    
    POS(i).Value = "(" & i + 1 & ") " & Playernames(i)
    
    Psheet.Cells(i + 1, 1) = POS(i)
    
    POS(i).HorizontalAlignment = xlCenter
    
Next i
  

End Sub


Private Sub AElement_Click()

UserForm1.Show

End Sub



Private Sub Damage_Click()

Dim Numeral2 As Long
Dim NameF As Long
Dim Optionlist() As Variant

Dim DamageFrm As UserForm2

Dim POS() As Variant
Dim Optbtn() As Variant
Dim Cell As Range
Dim i As Integer
Dim TopVar As Integer

Dim Playernames() As Variant

Dim Psheet As Worksheet

Set Psheet = ThisWorkbook.Worksheets("PlayerSheet")

Numeral2 = Psheet.Range("A1").End(xlDown).Row

ReDim Playernames(Numeral2 - 1)
ReDim POS(Numeral2 - 1)
ReDim Optionlist(Numeral2 - 1)
ReDim Optbtn(Numeral2 - 1)

Application.ScreenUpdating = False

For i = 1 To Numeral2

    Playernames(i - 1) = Psheet.Cells(i, 1).Value
    
Next i

'Need to come up with a variable range solution for my userform options - for now - 1 to 5 static'

'For i = 0 To numeral2 - 1
'    For Each Cell In ThisWorkbook.Worksheets("battlesheet").Range("A1:AW50")
'
'        If Cell.Value Like "(*" Then
'            PlayerNames(i) = Cell.Value
'            POS(i) = Cell.Address
'        End If
'
'Next Cell


        
TopVar = 10

For i = 0 To Numeral2 - 1
    Set Optbtn(i) = UserForm2.Controls("OptionButton" & i + 1)
    Optbtn(i).Caption = Playernames(i)
    Optbtn(i).Left = 15
    Optbtn(i).Top = TopVar
    Optbtn(i).Visible = True
    TopVar = TopVar + 15

Next i

Dim cntrl As control
For Each cntrl In UserForm2.Controls
    If cntrl.Caption Like "Option*" Then
        cntrl.Visible = False
    End If
Next cntrl

UserForm2.Caption = "Select Player to Damage / Heal:"

UserForm2.Show

End Sub

Private Sub DealENEDamage_Click()



Dim Numeral2 As Long
Dim NameF As Long
Dim Optionlist() As Variant

Dim DamageFrm As UserForm2

Dim POS() As Variant
Dim Optbtn() As Variant
Dim Cell As Range
Dim i As Integer
Dim TopVar As Integer

Dim Playernames() As Variant

Dim Psheet As Worksheet

Set Psheet = ThisWorkbook.Worksheets("PlayerSheet")

Numeral2 = Psheet.Range("C1").End(xlDown).Row

ReDim Playernames(Numeral2 - 1)
ReDim POS(Numeral2 - 1)
ReDim Optionlist(Numeral2 - 1)
ReDim Optbtn(Numeral2 - 1)

Application.ScreenUpdating = False

For i = 1 To Numeral2

    Playernames(i - 1) = Psheet.Cells(i, 3).Value
    
Next i
        
TopVar = 25

For i = 0 To Numeral2 - 1
    Set Optbtn(i) = UserForm4.Controls("Enemy" & i + 1)
    Optbtn(i).Caption = Playernames(i)
    Optbtn(i).Left = 15
    Optbtn(i).Top = TopVar
    Optbtn(i).Visible = True
    TopVar = TopVar + 20

Next i

Dim cntrl As control
For Each cntrl In UserForm4.Controls
    If cntrl.Caption Like "Option*" Then
        cntrl.Visible = False
    End If
Next cntrl

UserForm4.Caption = "Select Enemy to Damage:"


UserForm4.Show

End Sub

Private Sub MoveDown_Click()

Application.ScreenUpdating = False

Dim Psheet As Worksheet

Dim ptA As Range
Dim ptB As Range
Dim ptD As Range
Dim ptC As Range
Dim Bsheet As Worksheet

Set Bsheet = ThisWorkbook.Worksheets("BattleSheet")

Dim MoveRNG As Range

Set ptA = ActiveCell
Set ptB = ptA.Offset(1, 0)
Set ptC = Bsheet.Cells(53, 1)
Set ptD = Bsheet.Cells(53, 3)

ptD.Copy ptC

ptB.Copy ptD

ptA.Copy ptB

ptC.Copy ptA



ptA.ClearContents

ptB.Select

Set MoveRNG = Bsheet.Range("B2:AW50")


End Sub

Private Sub MoveLeft_Click()

Application.ScreenUpdating = False

Dim Psheet As Worksheet

Dim ptA As Range
Dim ptB As Range
Dim ptD As Range
Dim ptC As Range
Dim Bsheet As Worksheet

Set Bsheet = ThisWorkbook.Worksheets("BattleSheet")

Dim MoveRNG As Range

Set ptA = ActiveCell
Set ptB = ptA.Offset(0, -1)
Set ptC = Bsheet.Cells(53, 1)
Set ptD = Bsheet.Cells(53, 3)

ptD.Copy ptC

ptB.Copy ptD

ptA.Copy ptB

ptC.Copy ptA



ptA.ClearContents

ptB.Select

Set MoveRNG = Bsheet.Range("B2:AW50")

End Sub

Private Sub MoveRight_Click()

Application.ScreenUpdating = False

Dim Psheet As Worksheet

Dim ptA As Range
Dim ptB As Range
Dim ptD As Range
Dim ptC As Range
Dim Bsheet As Worksheet

Set Bsheet = ThisWorkbook.Worksheets("BattleSheet")

Dim MoveRNG As Range

Set ptA = ActiveCell
Set ptB = ptA.Offset(0, 1)
Set ptC = Bsheet.Cells(53, 1)
Set ptD = Bsheet.Cells(53, 3)

ptD.Copy ptC

ptB.Copy ptD

ptA.Copy ptB

ptC.Copy ptA



ptA.ClearContents

ptB.Select

Set MoveRNG = Bsheet.Range("B2:AW50")

End Sub

Private Sub MoveUp_Click()

Application.ScreenUpdating = False

Dim Psheet As Worksheet

Dim ptA As Range
Dim ptB As Range
Dim ptD As Range
Dim ptC As Range
Dim Bsheet As Worksheet

Set Bsheet = ThisWorkbook.Worksheets("BattleSheet")

Dim MoveRNG As Range

Set ptA = ActiveCell
Set ptB = ptA.Offset(-1, 0)
Set ptC = Bsheet.Cells(53, 1)
Set ptD = Bsheet.Cells(53, 3)

ptD.Copy ptC

ptB.Copy ptD

ptA.Copy ptB

ptC.Copy ptA

ptA.ClearContents

ptB.Select

Set MoveRNG = Bsheet.Range("B2:AW50")
End Sub

Private Sub REMOVEALL_Click()

Dim Numeral As Range

Set Numeral = ThisWorkbook.Worksheets("Battlesheet").Range("B2:AW50")
Numeral.ClearContents
        
Dim Numeral2 As Range

Set Numeral2 = ThisWorkbook.Worksheets("Playersheet").Range("A:C")
Numeral2.ClearContents
        
Dim Cell As Range

Application.ScreenUpdating = False

For Each Cell In Numeral
    
    Cell.Interior.Color = RGB(255, 255, 183)
    Cell.Interior.Pattern = xlPatternGray16
    Cell.Interior.PatternColor = RGB(204, 153, 0)
    Cell.Borders.LineStyle = xlDot
    
Next Cell


UserForm1.AElementS.Value = False

End Sub

Private Sub RemoveElements_Click()

UserForm3.Show

End Sub

Private Sub REnemies_Click()

Dim ENEArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Set Numeral = Application.InputBox( _
        Title:="Removing Enemies", _
        Prompt:="Select Area to Remove Enemies:", _
        Type:=8)
        
Dim Cell As Range

Application.ScreenUpdating = False

For Each Cell In Numeral
    If Cell.Interior.Color = RGB(200, 0, 0) _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(Cell.Value).Delete
        Cell.ClearContents
    End If
Next Cell

End Sub

Private Sub RPlayers_Click()

Dim ENEArea() As Variant
Dim Numeral As Variant
Dim Numerals4u() As Variant
Dim RType As Variant
Dim Ctrl As control

Dim ENELIST As Range

Set ENELIST = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C")

Set Numeral = Application.InputBox( _
        Title:="Removing Friendlies", _
        Prompt:="Select Area to Remove Friendlies:", _
        Type:=8)
        
Dim Cell As Range

Application.ScreenUpdating = False

For Each Cell In Numeral
    If Cell.Value Like "(*" _
    Then
        Cell.Interior.Color = RGB(255, 255, 183)
        Cell.Interior.Pattern = xlPatternGray16
        Cell.Interior.PatternColor = RGB(204, 153, 0)
        ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(Cell.Value).Delete
        Cell.ClearContents
    End If
Next Cell

End Sub
