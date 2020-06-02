VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptionButton1_Click()

UserForm2.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(UserForm2.OptionButton1.Caption)

Set RekPOS = ThisWorkbook.Worksheets("BattleSheet").Range("B2:AW50").Find(RektP.Value)

'Debug.Print RektP.Value

DMGTake = InputBox("Change in HP(-DMG , +HP):", "Modifying Player HP")

DMGDealt = Trim(Right(RektP.Value, 4))

If DMGDealt = "" Then
    Str = CStr(CInt(DMGTake) + 0)
    StrLEN = 4
Else
    Str = CStr(CInt(DMGTake) + CInt(DMGDealt))
    StrLEN = Len(DMGDealt)
End If


If DMGDealt = "" Then

    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - 4) & "  " & Str
    RekPOS.Value = Left(RekPOS, Len(RekPOS.Value) - 4) & "  " & Str
Else
    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str
    RekPOS.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str

End If

UserForm2.OptionButton1.Value = False

End Sub

Private Sub OptionButton2_Click()


UserForm2.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(UserForm2.OptionButton2.Caption)

Set RekPOS = ThisWorkbook.Worksheets("BattleSheet").Range("B2:AW50").Find(RektP.Value)

'Debug.Print RektP.Value

DMGTake = InputBox("Change in HP(-DMG , +HP):", "Modifying Enemy HP")

DMGDealt = Trim(Right(RektP.Value, 4))

If DMGDealt = "" Then
    Str = CStr(CInt(DMGTake) + 0)
    StrLEN = 4
Else
    Str = CStr(CInt(DMGTake) + CInt(DMGDealt))
    StrLEN = Len(DMGDealt)
End If


If DMGDealt = "" Then

    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - 4) & "  " & Str
    RekPOS.Value = Left(RekPOS, Len(RekPOS.Value) - 4) & "  " & Str
Else
    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str
    RekPOS.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str

End If

UserForm2.OptionButton2.Value = False

End Sub

Private Sub OptionButton3_Click()

UserForm2.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(UserForm2.OptionButton3.Caption)

Set RekPOS = ThisWorkbook.Worksheets("BattleSheet").Range("B2:AW50").Find(RektP.Value)

'Debug.Print RektP.Value

DMGTake = InputBox("How Much Damage (- integer):", "Dealing Damage")

DMGDealt = Trim(Right(RektP.Value, 4))

If DMGDealt = "" Then
    Str = CStr(CInt(DMGTake) + 0)
    StrLEN = 4
Else
    Str = CStr(CInt(DMGTake) + CInt(DMGDealt))
    StrLEN = Len(DMGDealt)
End If


If DMGDealt = "" Then

    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - 4) & "  " & Str
    RekPOS.Value = Left(RekPOS, Len(RekPOS.Value) - 4) & "  " & Str
Else
    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str
    RekPOS.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str

End If

UserForm2.OptionButton3.Value = False

End Sub

Private Sub OptionButton4_Click()

UserForm2.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(UserForm2.OptionButton4.Caption)

Set RekPOS = ThisWorkbook.Worksheets("BattleSheet").Range("B2:AW50").Find(RektP.Value)

'Debug.Print RektP.Value

DMGTake = InputBox("How Much Damage (- integer):", "Dealing Damage")

DMGDealt = Trim(Right(RektP.Value, 4))

If DMGDealt = "" Then
    Str = CStr(CInt(DMGTake) + 0)
    StrLEN = 4
Else
    Str = CStr(CInt(DMGTake) + CInt(DMGDealt))
    StrLEN = Len(DMGDealt)
End If


If DMGDealt = "" Then

    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - 4) & "  " & Str
    RekPOS.Value = Left(RekPOS, Len(RekPOS.Value) - 4) & "  " & Str
Else
    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str
    RekPOS.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str

End If

UserForm2.OptionButton4.Value = False

End Sub

Private Sub OptionButton5_Click()
UserForm2.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("A:A").Find(UserForm2.OptionButton5.Caption)

Set RekPOS = ThisWorkbook.Worksheets("BattleSheet").Range("B2:AW50").Find(RektP.Value)

'Debug.Print RektP.Value

DMGTake = InputBox("How Much Damage (- integer):", "Dealing Damage")

DMGDealt = Trim(Right(RektP.Value, 4))

If DMGDealt = "" Then
    Str = CStr(CInt(DMGTake) + 0)
    StrLEN = 4
Else
    Str = CStr(CInt(DMGTake) + CInt(DMGDealt))
    StrLEN = Len(DMGDealt)
End If


If DMGDealt = "" Then

    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - 4) & "  " & Str
    RekPOS.Value = Left(RekPOS, Len(RekPOS.Value) - 4) & "  " & Str
Else
    RektP.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str
    RekPOS.Value = Left(RekPOS.Value, Len(RekPOS.Value) - StrLEN - 2) & "  " & Str

End If

UserForm2.OptionButton5.Value = False

End Sub
