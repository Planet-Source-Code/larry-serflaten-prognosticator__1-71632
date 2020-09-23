VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum SS
  max = 0
  Value
  Owned
  Available
  Trigger
End Enum

Private WithEvents Buy As CommandButton
Attribute Buy.VB_VarHelpID = -1
Private WithEvents Sell As CommandButton
Attribute Sell.VB_VarHelpID = -1
Private WithEvents Clock As Timer
Attribute Clock.VB_VarHelpID = -1
Private WithEvents Chart As PictureBox
Attribute Chart.VB_VarHelpID = -1

Private Stocks As Tripod
Private Reports As Collection
Private Palette(1 To 27) As Long
Private Stats(1 To 27, 0 To 4) As Currency

Private Selected As Long
Private TimeLine As Long
Private MouseClick As Long


Public Sub Setup()
Dim i, CTLS
  ' Grab control references
  Set CTLS = Main.Controls
  Set Buy = CTLS("BTN1")
  Set Sell = CTLS("BTN2")
  Set Clock = CTLS("TMR1")
  Set Chart = CTLS("PIC1")
  ' Build color palette
  For i = 0 To 26
    Palette(i + 1) = CTLS("PAL1").Point(i, 0)
  Next
  Cash = Bankroll
  ' Build Form
  GameScreen
End Sub

Public Sub Execute()
Dim i
  ' Build stock entities
  Set Stocks = New Tripod
  Set Reports = New Collection
  Stocks.Target = 100
  Stocks.Value = 10
  Stocks.Build 2
  
  ' Advance from initial state
  Stocks.Target = 10000
  For i = 1 To 50
    Stocks.Tick (Rnd < 0.3)
  Next
  
  ' Init statistics
  For i = 1 To 27
    Stats(i, SS.max) = 2000
    Stats(i, SS.Owned) = 0
    Stats(i, SS.Trigger) = 1
    Stats(i, SS.Value) = 0
    Stats(i, SS.Available) = 500
  Next

  TimeLine = 7390
  Main.Line (270, 90)-Step(TimeLine, 90), vb3DLight, BF
  TimeLine = TimeLine + TimeLineSpeed
  
  ' Set up timer
  Clock.Interval = GameTimerInterval
  Clock.Enabled = True
  Clock_Timer
End Sub

Private Sub GameScreen()
  ' Graphics
  Set Main.Picture = Nothing
  Chart.Top = 300
  Chart.Height = 3000
  Chart.Scale (-1, 1800)-(30, -180)
  Set Chart.Picture = Nothing
  Chart.BackColor = vbActiveTitleBar
  Chart.ForeColor = vb3DHighlight
  Chart.Font.Name = "Symbol"
  Chart.Font.Size = 12
  With Main.Font
    .Name = "Tahoma"
    .Size = 10
    .Bold = False
  End With
  Buy.Move 4400, 3400, 1200, 390
  Buy.Caption = "&Buy"
  Set Buy.Font = FontClone(Main.Font)
  Buy.Visible = True
  Sell.Move 6465, 3400, 1200, 390
  Sell.Caption = "&Sell"
  Set Sell.Font = FontClone(Main.Font)
  Sell.Visible = True
  Main.FillStyle = vbFSSolid
  Main.FillColor = vbWhite
  Main.Line (270, 3900)-Step(4000, 1500), vbBlack, B
  ' Stock chart graphics
  Main.ForeColor = vbBlack
  Main.PSet (300, 4100), vbWhite
  Main.Print "200"
  Main.PSet (300, 4600), vbWhite
  Main.Print "100"
  Main.PSet (510, 5100), vbWhite
  Main.Print "0"
  Main.Line (630, 3930)-Step(0, 1400), vbBlack
  Main.Line -Step(3500, 0), vbBlack
  Main.Line (645, 3930)-Step(0, 1380), vbBlack
  Main.Line -Step(3500, 0), vbBlack
  ' Stock Report graphics
  Main.FillColor = vbApplicationWorkspace
  Main.Line (4400, 3900)-Step(3255, 1500), vbBlack, B
  Main.PSet (4490, 3945), vbApplicationWorkspace
  Main.Print "STOCK CAP"
  Main.PSet (4490, 4190), vbApplicationWorkspace
  Main.Print "STOCK VALUE"
  Main.PSet (4490, 4515), vbApplicationWorkspace
  Main.Print "EQUITY OWNED"
  Main.PSet (4490, 4760), vbApplicationWorkspace
  Main.Print "EQUITY TOTAL"
  Main.PSet (4490, 5100), vbApplicationWorkspace
  Main.Print "AVAILABLE"
  UpdateCash
End Sub

Private Sub UpdateCash()
Dim TXT As String
  Main.Line (300, 3390)-Step(4000, 480), vbApplicationWorkspace, BF
  With Main.Font
    .Name = "Tahoma"
    .Size = 18
    .Bold = True
  End With
  TXT = Format$(Cash, "currency")
  Main.ForeColor = vb3DDKShadow
  Main.PSet (4300 - Main.TextWidth(TXT), 3410), vbApplicationWorkspace
  Main.Print TXT
  Main.ForeColor = vb3DHighlight
  Main.PSet (4285 - Main.TextWidth(TXT), 3390), vbApplicationWorkspace
  Main.Print TXT
  If Cash <= 99999999.99 Then
  Main.ForeColor = vb3DDKShadow
  Main.PSet (285, 3410), vbApplicationWorkspace
  Main.Print "CASH"
  Main.ForeColor = vb3DHighlight
  Main.PSet (270, 3390), vbApplicationWorkspace
  Main.Print "CASH"
  End If
  With Main.Font
    .Name = "Tahoma"
    .Size = 10
    .Bold = False
  End With
End Sub

Private Sub UpdateReport()
Dim i
  ' Text report
  If Stats(Selected, SS.Owned) > 0 Then
    Stats(Selected, SS.Available) = Int((Stats(Selected, SS.max) - Stats(Selected, SS.Owned)) * 0.07) * 10
  End If
  Main.Line (5900, 3930)-Step(1690, 1400), vbApplicationWorkspace, BF
  AlignText 3945, CStr(Stats(Selected, SS.max))
  AlignText 4190, Format(Stats(Selected, SS.Value), "0.00")
  AlignText 4515, CStr(Stats(Selected, SS.Owned))
  AlignText 4760, Format(Stats(Selected, SS.Value) * Stats(Selected, SS.Owned), "#,##0.00")
  AlignText 5100, CStr(Stats(Selected, SS.Available))
  ' Graph report
  Main.Line (700, 3930)-Step(3560, 1360), vbWhite, BF
  Main.Line (700, 3930)-Step(180, 180), Palette(Selected), BF
  Main.PSet (700, 5300 - Reports(1)(Selected) / 2), vbWhite
  For i = 1 To Reports.Count
    Main.Line -(i * 75 + 630, 5300 - Reports(i)(Selected) / 2), vbBlack
  Next
End Sub

Private Sub UpdateChart()
Dim i As Long, V
  Set V = Reports(Reports.Count)
  Equity = 0
  For i = 1 To 27
    ' Bankrupt
    If V(i) <= 0 Then
       Stats(i, SS.Trigger) = 1
       If Stats(i, SS.Owned) > 0 Then
         Beep
         Stats(i, SS.Owned) = 0
         Stats(i, SS.Available) = 500
         Chart.Line (i, -2)-Step(0.9, -200), vbActiveTitleBar, BF
       End If
    End If
    
    ' Stock split
    If V(1) > 2000 Then
      If Stats(i, SS.Trigger) = 1 Then
        Stats(i, SS.max) = Stats(i, SS.max) * 2
        Stats(i, SS.Trigger) = -1
      End If
    End If
    
    Stats(i, SS.Value) = V(i) / 10
    Equity = Equity + Stats(i, SS.Value)
    Chart.Line (i, 0)-Step(1, V(i)), Palette(i), BF
    Chart.Line (i, V(i))-(i + 1, Chart.ScaleTop), vbActiveTitleBar, BF
  Next
  If Selected > 0 Then
    Buy.Enabled = (Stats(Selected, SS.Available) > 0) And (Stats(Selected, SS.Value) > 0)
    Sell.Enabled = (Stats(Selected, SS.Owned) > 0)
  End If
End Sub

Private Sub AlignText(ByVal Y As Long, ByVal TXT As String)
  Main.PSet (7600 - Main.TextWidth(TXT), Y), vbApplicationWorkspace
  Main.Print TXT
End Sub

Private Sub Buy_Click()
Dim max As Long, amt As Long
  If Selected = 0 Then Exit Sub
  Clock.Enabled = False
  max = Stats(Selected, SS.Available)
  If Cash < (max * (Stats(Selected, SS.Value))) Then
     max = Int(Cash / Stats(Selected, SS.Value))
  End If
  amt = Trade.Amount(max)
  Set Trade = Nothing
  Cash = Cash - (Stats(Selected, SS.Value) * amt)
  Stats(Selected, SS.Owned) = Stats(Selected, SS.Owned) + amt
  Stats(Selected, SS.Available) = Int((Stats(Selected, SS.max) - Stats(Selected, SS.Owned)) * 0.07) * 10
  UpdateCash
  Buy.Enabled = (Stats(Selected, SS.Available) > 0) And (Stats(Selected, SS.Value) > 0)
  Sell.Enabled = (Stats(Selected, SS.Owned) > 0)
  Clock.Enabled = True
End Sub

Private Sub Chart_DblClick()
  Select Case MouseClick
  Case vbLeftButton
    If Buy.Enabled Then Buy_Click
  Case vbRightButton
    If Sell.Enabled Then Sell_Click
  End Select
End Sub

Private Sub Sell_Click()
Dim amt As Long
  Clock.Enabled = False
  amt = Trade.Amount(Stats(Selected, SS.Owned))
  Set Trade = Nothing
  Cash = Cash + (Stats(Selected, SS.Value) * amt)
  Stats(Selected, SS.Owned) = Stats(Selected, SS.Owned) - amt
  Stats(Selected, SS.Available) = Int((Stats(Selected, SS.max) - Stats(Selected, SS.Owned)) * 0.07) * 10
  UpdateCash
  Buy.Enabled = (Stats(Selected, SS.Available) > 0) And (Stats(Selected, SS.Value) > 0)
  Sell.Enabled = (Stats(Selected, SS.Owned) > 0)
  Clock.Enabled = True
End Sub

Private Sub Chart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const Mark As String = "¨"
   MouseClick = Button
   If (Int(X) > 0) And (Int(X) < 28) Then
      Chart.Line (Selected, -2)-Step(0.9, -200), vbActiveTitleBar, BF
      If Selected > 0 Then
         If Stats(Selected, SS.Owned) > 0 Then
            Chart.Line (Selected, -2)-Step(0.9, -200), vbApplicationWorkspace, BF
         End If
      End If
      Selected = Int(X)
      Chart.PSet ((Selected + 0.5) - Chart.TextWidth(Mark) / 2, -1), vbActiveTitleBar
      Chart.Print Mark
      UpdateReport
      Buy.Enabled = (Stats(Selected, SS.Available) > 0) And (Cash >= Stats(Selected, SS.Available)) And (Stats(Selected, SS.Value) > 0)
      Sell.Enabled = Stats(Selected, SS.Owned) > 0
   End If
End Sub

Private Sub Chart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Chart_MouseDown Button, Shift, X, Y
   If Button = vbRightButton Then Clock.Enabled = Not Clock.Enabled
End Sub

Private Sub Clock_Timer()
Dim V As New Collection
Dim i

  ' Process time line
  TimeLine = TimeLine - TimeLineSpeed
  Main.Line (270 + TimeLine, 90)-Step(0, 90), vb3DShadow, BF
  If TimeLine <= 0 Then
    Clock.Enabled = False
    Main.StateChange New Scoring
    Exit Sub
  End If
  
  ' Occasional target adjustments
  If (Int(Rnd * 100) Mod 7) = 0 Then
    Stocks.Target = 20000
    Stocks.Tick True
    Stocks.Tick True
  ElseIf (Int(Rnd * 100) Mod 17) = 0 Then
    Stocks.Target = -5000
    Stocks.Tick True
    Stocks.Tick True
    Stocks.Tick True
  ElseIf (Int(Rnd * 100) Mod 5) = 0 Then
    Stocks.Target = Stocks.Target + 100
    Stocks.Tick True
  Else
    Stocks.Tick False
  End If
  ' Get stock values
  Stocks.Tally V
  Reports.Add V
  If Reports.Count > 45 Then Reports.Remove 1
  ' Display stock values
  UpdateChart
  If Selected > 0 Then UpdateReport
End Sub

Private Function FontClone(Fnt As StdFont) As StdFont
  Set FontClone = New StdFont
  With FontClone
    .Name = Fnt.Name
    .Bold = Fnt.Bold
    .Size = Fnt.Size
  End With
End Function
