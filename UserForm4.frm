VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Enemy1_Click()

UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy1.Caption)

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

UserForm4.Enemy1.Value = False


End Sub

Private Sub Enemy2_Click()

UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy2.Caption)

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

UserForm4.Enemy2.Value = False

End Sub

Private Sub Enemy3_Click()

UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy3.Caption)

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

UserForm4.Enemy3.Value = False

End Sub

Private Sub Enemy4_Click()
UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy4.Caption)

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

UserForm4.Enemy4.Value = False
End Sub

Private Sub Enemy5_Click()
UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy5.Caption)

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

UserForm4.Enemy5.Value = False

End Sub

Private Sub Enemy6_Click()
UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy6.Caption)

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

UserForm4.Enemy6.Value = False
End Sub

Private Sub Enemy7_Click()
UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy7.Caption)

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

UserForm4.Enemy7.Value = False
End Sub

Private Sub Enemy8_Click()

UserForm4.Hide

Dim DMGTake As Variant
Dim DMGDealt As String
Dim RektP As Range
Dim RekPOS As Range
Dim StrLEN As Integer
Dim Str As String

Set RektP = ThisWorkbook.Worksheets("PlayerSheet").Range("C:C").Find(UserForm4.Enemy8.Caption)

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

UserForm4.Enemy8.Value = False

End Sub
