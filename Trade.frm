VERSION 5.00
Begin VB.Form Trade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade"
   ClientHeight    =   1305
   ClientLeft      =   5730
   ClientTop       =   5505
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2970
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   50
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1740
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "01000"
      Height          =   255
      Left            =   2100
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   150
      Width           =   195
   End
End
Attribute VB_Name = "Trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Value As Long

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  Value = 0
  Unload Me
End Sub

Public Function Amount(ByVal max As Long) As Long
  HScroll1.max = max
  Label2.Caption = max
  Value = Int(max / 2)
  HScroll1.Value = Value
  Me.Move Main.Left + 4400, Main.Top + 1800
  Me.Show vbModal
  Amount = Value
End Function

Private Sub HScroll1_Change()
  Text1.Text = CStr(HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
  Text1.Text = CStr(HScroll1.Value)
End Sub

Private Sub Text1_Change()
  Value = Val(Text1)
  If Value > HScroll1.max Then Value = HScroll1.max
  Text1.Text = CStr(Value)
End Sub
