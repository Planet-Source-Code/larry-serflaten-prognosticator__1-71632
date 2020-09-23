VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prognosticator"
   ClientHeight    =   4650
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4785
   Begin VB.CommandButton BTN2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3390
      Width           =   1635
   End
   Begin VB.PictureBox PAL1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      Picture         =   "Main.frx":0000
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton BTN1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   240
      TabIndex        =   1
      Top             =   2580
      Width           =   1635
   End
   Begin VB.Timer TMR1 
      Left            =   60
      Top             =   150
   End
   Begin VB.PictureBox PIC1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   180
      ScaleHeight     =   1665
      ScaleWidth      =   2205
      TabIndex        =   0
      Top             =   780
      Width           =   2235
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private State As Object

Private Sub Form_Load()
   StateChange New TitleScreen
End Sub

Public Sub StateChange(NewState As Object)
   Me.Hide
   Set State = NewState
   State.Setup
   Me.Show
   State.Execute
End Sub

Private Sub TMR1_Timer()
  TMR1.Enabled = (Me.WindowState = vbNormal)
End Sub
