VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tripod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Target As Currency
Public Value As Currency
Private Legs(0 To 2) As Tripod

Public Sub Tick(ByVal Change As Boolean)
Static fraction As Single
Static delta As Currency
  If Change Then
    If ((Int(Rnd * 100) Mod 3) = 0) Then
      fraction = (Rnd / 9) + 0.002
      delta = Target - Value
    Else
      fraction = fraction * 0.8
    End If
  End If
  Value = Value + (fraction * delta)
'Debug.Print "T:"; Target, "V:"; Value
  
  If Not Legs(0) Is Nothing Then Triplette Change

End Sub

Public Sub Triplette(Change As Boolean)
Dim idx, itm, mix As New Collection

   If Change Then
      itm = (Rnd / 3) + 0.2
      mix.Add Value * itm
      itm = (Rnd / 3) + 0.2
      mix.Add Value * itm
      mix.Add Value - (mix(1) + mix(2))
      ' shuffle
      For itm = 1 To 5
        idx = Int(Rnd * 3) + 1
        mix.Add mix(idx)
        mix.Remove idx
      Next
      For idx = 0 To 2
        Legs(idx).Target = mix(idx + 1)
      Next
   End If
 
   For idx = 0 To 2
     Legs(idx).Tick Change
   Next
End Sub

Public Sub Build(Optional Level As Long)
Dim idx As Long
  For idx = 0 To 2
   Set Legs(idx) = New Tripod
   If Level > 0 Then
    Legs(idx).Target = Target / (2.5 + Rnd)
    Legs(idx).Value = Value
    Legs(idx).Build Level - 1
   End If
  Next
End Sub

Public Sub Tally(Values As Collection)
Dim idx As Long
  If Legs(0) Is Nothing Then
    If Value > 0 Then
      Values.Add Value
    Else
      Values.Add 0
    End If
  Else
    For idx = 0 To 2
      Legs(idx).Tally Values
    Next
  End If
End Sub



