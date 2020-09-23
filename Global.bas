Attribute VB_Name = "Global"
Option Explicit

Public Cash As Currency
Public Equity As Currency
Public High As Currency

' Values you might want to tweak... (Stay within limits for best results)

Public Const GameTimerInterval = 200  ' Chart speed adjustment (10 - 2000)
Public Const TimeLineSpeed = 5        ' Game duration adjustment (1 - 15)
Public Const Bankroll = 5000          ' Initial Cash ( 1 - 99999 )

