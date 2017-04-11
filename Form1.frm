VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Const PI = 3.141592653

Private xCenter As Long, yCenter As Long
Private WithEvents HotKeys As cHotKey
Attribute HotKeys.VB_VarHelpID = -1
Private R As Long
Private fs As New CFile
Private FilePath As String


Private Sub Form_Load()
  FilePath = App.Path & "\runlog.txt"
  R = (Screen.Height / Screen.TwipsPerPixelY / 2) - 5
  xCenter = Screen.Width / Screen.TwipsPerPixelX / 2
  yCenter = Screen.Height / Screen.TwipsPerPixelY / 2
  Set HotKeys = New cHotKey
  HotKeys.AddHotKey vbKeyE, True
  HotKeys.StartHotKeys Timer2
  Log "程序启动：" & Format(Now, "yyyy-MM-dd HH:mm:ss")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  HotKeys.StopHotKeys
  HotKeys.ClsAllHotKey
End Sub

Private Sub HotKeys_HotKeyPress(ByVal HotKey As Long, ByVal hCtrl As Boolean, ByVal hAlt As Boolean, ByVal hShift As Boolean)
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Dim x As Long, y As Long, sec As Long
  sec = Second(Now)
  x = xCenter + R * Sin(PI * sec / 30)
  y = yCenter - R * Cos(PI * sec / 30)
  SetCursorPos x, y
End Sub

Private Sub Timer3_Timer()
  Static count As Long
  count = count + 1
  If count >= 5 Then
    Log Format(Now, "yyyy-MM-dd HH:mm:ss")
    count = 0
  End If
  
End Sub

Private Sub Log(ByVal Message As String)
  fs.WriteLineToTextFile FilePath, Message
End Sub
