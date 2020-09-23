VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Scoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents OK As CommandButton
Attribute OK.VB_VarHelpID = -1
Private WithEvents Quit As CommandButton
Attribute Quit.VB_VarHelpID = -1



Public Sub Setup()
Dim ctl
  On Error Resume Next
  For Each ctl In Main.Controls
    ctl.Visible = False
  Next
  On Error GoTo 0
  
  Set Main.Picture = Nothing
  Set OK = Main.Controls("BTN1")
  Set Quit = Main.Controls("BTN2")
  OK.Move 270, 5000
  OK.Visible = True
  OK.Caption = "Play Again"
  OK.Enabled = True
  Quit.Move 6500, 5000
  Quit.Visible = True
  Quit.Caption = "Exit"
  Quit.Enabled = True
  Main.Font.Name = "Tahoma"
  Main.Font.Size = 36
  Main.Font.Bold = True
  LAlign 390, 0, "PROGNOSTICATOR!"
  Main.Font.Size = 18
  LAlign 600, 1400, "CASH"
  RAlign 7400, 1400, Format(Cash, "#,##0.00")
  LAlign 600, 2000, "EQUITY"
  RAlign 7400, 2000, Format(Equity, "#,##0.00")
  Main.Font.Size = 24
  LAlign 600, 3200, "TOTAL"
  RAlign 7400, 3200, Format(Cash + Equity, "0,000.00")
  Main.Font.Size = 12
  Main.ForeColor = vb3DHighlight
  If (Equity + Cash) > High Then
    Main.PSet (880, 4400), vbApplicationWorkspace
    Main.Print "CONGRATULATIONS ON YOUR NEW HIGH SCORE!"
    High = Equity + Cash
  Else
    Main.PSet (2000, 4400), vbApplicationWorkspace
    Main.Print "HIGH SCORE = "; Format(High, "#,##0.00")
  End If
  
  
End Sub
Public Sub Execute()
'
End Sub

Private Sub OK_Click()
  Main.StateChange New TitleScreen
End Sub

Private Sub Quit_Click()
  Unload Main
End Sub

Private Sub RAlign(ByVal X As Long, ByVal Y As Long, TXT As String)

  Main.ForeColor = vb3DDKShadow
  Main.PSet (X + 30 - Main.TextWidth(TXT), Y + 30), vbApplicationWorkspace
  Main.Print TXT
  Main.ForeColor = vb3DHighlight
  Main.PSet (X - Main.TextWidth(TXT), Y), vbApplicationWorkspace
  Main.Print TXT
End Sub

Private Sub LAlign(ByVal X As Long, ByVal Y As Long, TXT As String)

  Main.ForeColor = vb3DDKShadow
  Main.PSet (X + 30, Y + 30), vbApplicationWorkspace
  Main.Print TXT
  Main.ForeColor = vb3DHighlight
  Main.PSet (X, Y), vbApplicationWorkspace
  Main.Print TXT
  
End Sub
