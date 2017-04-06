VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'**模 块 名：cHotKey
'**说    明：定义全局热键
'**创 建 人：马大哈 http://www.m5home.com/
'**日    期：2006年7月19日
'**版    本：V1.0
'*************************************************************************
Option Explicit
Public Enum vbKeyAll                    '枚举值,可根据自己的实际情况调整
'    vbKeyLButton = 1 '鼠标左键
'    vbKeyRButton = 2 '鼠标右键
'    vbKeyCancel = 3 'CANCEL 键
'    vbKeyMButton = 4 '鼠标中键
'    vbKeyBack = 8 'BACKSPACE 键
'    vbKeyTab = 9 'TAB 键
'    vbKeyClear = 12 'CLEAR 键
'    vbKeyReturn = 13 ' ENTER 键
'    vbKeyShift = 16 'SHIFT 键
'    vbKeyControl = 17 'CTRL 键
'    vbKeyMenu = 18 ' 菜单键
'    vbKeyPause = 19 'PAUSE 键
'    vbKeyCapital = 20 'CAPS LOCK 键
'    vbKeyEscape = 27 'ESC 键
'    vbKeySpace = 32 'SPACEBAR 键
'    vbKeyPageUp = 33 'PAGEUP 键
'    vbKeyPageDown = 34 'PAGEDOWN 键
'    vbKeyEnd = 35 'END 键
'    vbKeyHome = 36 'HOME 键
'    vbKeyLeft = 37 'LEFT ARROW 键
'    vbKeyUp = 38 'UP ARROW 键
'    vbKeyRight = 39 'RIGHT ARROW 键
'    vbKeyDown = 40 'DOWN ARROW 键
'    vbKeySelect = 41 'SELECT 键
'    vbKeyPrint = 42 'PRINT SCREEN 键
'    vbKeyExecute = 43 'EXECUTE 键
'    vbKeySnapshot = 44 'SNAP SHOT 键
'    vbKeyInser = 45 'INS 键
'    vbKeyDelete = 46 'DEL 键
'    vbKeyHelp = 47 'HELP 键
'    vbKeyNumlock = 144 'NUM LOCK 键
   
    vbKeyA = 65 'A 键
    vbKeyB = 66 'B 键
    vbKeyC = 67 'C 键
    vbKeyD = 68 'D 键
    vbKeyE = 69 'E 键
    vbKeyF = 70 'F 键
    vbKeyG = 71 'G 键
    vbKeyH = 72 'H 键
    vbKeyI = 73 'I 键
    vbKeyJ = 74 'J 键
    vbKeyK = 75 'K 键
    vbKeyL = 76 'L 键
    vbKeyM = 77 'M 键
    vbKeyN = 78 'N 键
    vbKeyO = 79 'O 键
    vbKeyP = 80 'P 键
    vbKeyQ = 81 'Q 键
    vbKeyR = 82 'R 键
    vbKeyS = 83 'S 键
    vbKeyT = 84 'T 键
    vbKeyU = 85 'U 键
    vbKeyV = 86 'V 键
    vbKeyW = 87 'W 键
    vbKeyX = 88 'X 键
    vbKeyY = 89 'Y 键
    vbKeyZ = 90 'Z 键
   
'    vbKey0 = 48 '0 键
'    vbKey1 = 49 '1 键
'    vbKey2 = 50 '2 键
'    vbKey3 = 51 '3 键
'    vbKey4 = 52 '4 键
'    vbKey5 = 53 '5 键
'    vbKey6 = 54 '6 键
'    vbKey7 = 55 '7 键
'    vbKey8 = 56 '8 键
'    vbKey9 = 57 '9 键
'    vbKeyNumpad0 = 96 '0 键
'    vbKeyNumpad1 = 97 '1 键
'    vbKeyNumpad2 = 98 '2 键
'    vbKeyNumpad3 = 99 '3 键
'    vbKeyNumpad4 = 100 '4 键
'    vbKeyNumpad5 = 101 '5 键
'    vbKeyNumpad6 = 102 '6 键
'    vbKeyNumpad7 = 103 '7 键
'    vbKeyNumpad8 = 104 ' 8 键
'    vbKeyNumpad9 = 105 '9 键
'    vbKeyMultiply = 106 '乘号 (*) 键
'    vbKeyAdd = 107 '加号 (+) 键
'    vbKeySeparator = 108 'ENTER 键（在数字小键盘上）
'    vbKeySubtract = 109 '减号 (-) 键
'    vbKeyDecimal = 110 '小数点 (.) 键
'    vbKeyDivide = 111 '除号 (/) 键
   
    vbKeyF1 = 112 'F1 键
    vbKeyF2 = 113 'F2 键
    vbKeyF3 = 114 'F3 键
    vbKeyF4 = 115 'F4 键
    vbKeyF5 = 116 'F5 键
    vbKeyF6 = 117 'F6 键
    vbKeyF7 = 118 'F7 键
    vbKeyF8 = 119 'F8 键
    vbKeyF9 = 120 'F9 键
    vbKeyF10 = 121 'F10 键
    vbKeyF11 = 122 'F11 键
    vbKeyF12 = 123 'F12 键
    vbKeyF13 = 124 'F13 键
    vbKeyF14 = 125 'F14 键
    vbKeyF15 = 126 'F15 键
    vbKeyF16 = 127 'F16 键
End Enum
Private Const MOD_CONTROL = &H2
Private Const MOD_ALT = &H1
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312
   
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Event HotKeyPress(ByVal HotKey As Long, ByVal hCtrl As Boolean, ByVal hAlt As Boolean, ByVal hShift As Boolean)
Dim HotKeys() As Long           '热键值
Dim hCtrls() As Boolean         'Ctrl键状态
Dim hAlts() As Boolean          'Alt键状态
Dim hShifts() As Boolean        'Shift键状态
Dim KeyAtom() As Integer        '原子值
Dim WithEvents CtlTimer As Timer        '用于启动死循环的定时器
Attribute CtlTimer.VB_VarHelpID = -1
Dim bCancel As Boolean
Dim mfrmHwnd As Long           '主窗体句柄
Public Sub StartHotKeys(ByRef theTimer As Timer)
    '启动消息死循环,开始捕捉热键
    'theTimer:一个定时器
   
    bCancel = False
    Set CtlTimer = theTimer
    With CtlTimer
        .Enabled = False
        .Interval = 10
        .Tag = ""
        .Enabled = True
    End With
End Sub
Public Sub StopHotKeys()
    '停止消息循环,停止捕捉热键
   
    bCancel = True
    DoEvents
End Sub
Private Sub GetMessage()
    '消息循环,是一个死循环,需要由定时器调用,以便调用过程可以立即返回
    '在此循环内判断发送到目标句柄的热键消息,并以事件返回
   
    Dim Message As Msg
    Dim I As Long
   
    Do While Not bCancel
        WaitMessage
        If PeekMessage(Message, mfrmHwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            Debug.Print "HotKeyMessage: " & Message.Message
            For I = 1 To UBound(KeyAtom)            '有热键消息来时,判断一下是哪个热键,并返回其值
                If KeyAtom(I) = Message.wParam Then     'wParam里就是原子值
                    RaiseEvent HotKeyPress(HotKeys(I), hCtrls(I), hAlts(I), hShifts(I))
                    Exit For
                End If
            Next I
        End If
        DoEvents
    Loop
End Sub
Public Function AddHotKey(ByVal HotKey As vbKeyAll, Optional ByVal hCtrl As Boolean = False, Optional ByVal hAlt As Boolean = False, Optional ByVal hShift As Boolean = False) As Boolean
    '添加热键
    '如果返回值非零,则成功
    Dim I As Long
    Dim J As Long
    Dim K As Long
   
    AddHotKey = False
    I = InHotValue(HotKey)
   
    If I <> -1 Then      '如果按键已经有了,再判断一下控制键
        If hCtrls(I) = hCtrl And hShifts(I) = hShift Then       '如果控制键也相同,那就已经定义过了,退出
            Exit Function
        End If
    End If
   
    ReDim Preserve HotKeys(UBound(HotKeys) + 1)
    ReDim Preserve KeyAtom(UBound(KeyAtom) + 1)
    ReDim Preserve hCtrls(UBound(hCtrls) + 1)
    ReDim Preserve hAlts(UBound(hAlts) + 1)
    ReDim Preserve hShifts(UBound(hShifts) + 1)
   
    HotKeys(UBound(HotKeys)) = HotKey
    KeyAtom(UBound(KeyAtom)) = GlobalAddAtom(GetTmpStr(16))
    hCtrls(UBound(hCtrls)) = hCtrl
    hAlts(UBound(hAlts)) = hAlt
    hShifts(UBound(hShifts)) = hShift
   
    J = IIf(hCtrl = True, MOD_CONTROL, 0)       '组合控制键
    K = J
    J = IIf(hAlt = True, MOD_ALT, 0)
    K = IIf(J = 0, K, K Or J)
    J = IIf(hShift = True, MOD_SHIFT, 0)
    K = IIf(J = 0, K, K Or J)
   
    I = RegisterHotKey(mfrmHwnd, KeyAtom(UBound(KeyAtom)), K, HotKey)
    AddHotKey = I
   
    Debug.Print "HotKeyRegister:  " & I
   
    If I = 0 Then           '注册失败,就删除这个热键
        Call DelHotKey(HotKey, hCtrl, hAlt, hShift)
    End If
End Function
Public Function DelHotKey(ByVal HotKey As Long, Optional ByVal hCtrl As Boolean = False, Optional ByVal hAlt As Boolean = False, Optional ByVal hShift As Boolean = False) As Boolean
    '删除热键
    Dim I As Long
    Dim DelIndex As Long
   
    Dim tmpHotKeys() As Long                '交换用数组
    Dim tmpKeyAtom() As Integer             '交换用数组
    Dim tmphCtrls() As Boolean              '交换用数组
    Dim tmphAlts() As Boolean                '交换用数组
    Dim tmphShifts() As Boolean              '交换用数组
   
    On Error Resume Next
   
    DelHotKey = False
    DelIndex = InHotValue(HotKey, hCtrl, hAlt, hShift)        '取得索引
   
    If DelIndex = -1 Then Exit Function             '判断热键
    If hCtrl <> hCtrls(DelIndex) Then Exit Function
    If hAlt <> hAlts(DelIndex) Then Exit Function
    If hShift <> hShifts(DelIndex) Then Exit Function
   
    UnregisterHotKey mfrmHwnd, KeyAtom(DelIndex)             '先删除热键
    GlobalDeleteAtom KeyAtom(DelIndex)                      '再删除原子
   
    ReDim tmpHotKeys(I)           '初始化临时数组
    ReDim tmpKeyAtom(I)
    ReDim tmphCtrls(I)
    ReDim tmphAlts(I)
    ReDim tmphShifts(I)
   
    For I = 1 To UBound(HotKeys)            '把内容倒到临时数组内
        If I <> DelIndex Then
            ReDim tmpHotKeys(UBound(tmpHotKeys) + 1)
            ReDim tmpKeyAtom(UBound(tmpKeyAtom) + 1)
            ReDim tmphCtrls(UBound(tmphCtrls) + 1)
            ReDim tmphAlts(UBound(tmphAlts) + 1)
            ReDim tmphShifts(UBound(tmphShifts) + 1)
           
            tmpHotKeys(UBound(tmpHotKeys)) = HotKeys(I)
            tmpKeyAtom(UBound(tmpKeyAtom)) = KeyAtom(I)
            tmphCtrls(UBound(tmphCtrls)) = hCtrls(I)
            tmphAlts(UBound(tmphAlts)) = hAlts(I)
            tmphShifts(UBound(tmphShifts)) = hShifts(I)
        End If
    Next I
   
    I = UBound(tmpHotKeys)
   
    ReDim HotKeys(I)           '重定义数组大小,原内容不保存
    ReDim KeyAtom(I)
    ReDim hCtrls(I)
    ReDim hAlts(I)
    ReDim hShifts(I)
    HotKeys = tmpHotKeys         '再把调整后的内容倒回原数组
    KeyAtom = tmpKeyAtom
    hCtrls = tmphCtrls
    hAlts = tmphAlts
    hShifts = tmphShifts
   
    DelHotKey = True
End Function
Public Function ClsAllHotKey() As Boolean
    '删除所有热键.
    Dim I As Long
   
    On Error Resume Next
   
    bCancel = True
    DoEvents
   
    For I = 1 To UBound(HotKeys)
        Call DelHotKey(HotKeys(I), hCtrls(I), hAlts(I), hShifts(I))
    Next I
End Function
Private Sub Class_Initialize()
    '类初始化
    Call CleanData
    bCancel = True
End Sub
Private Sub Class_Terminate()
    '类销毁
    bCancel = True
    DoEvents
    Call ClsAllHotKey
    Call CleanData
End Sub
Private Sub CleanData()
    '清除数据
    ReDim HotKeys(0)
    ReDim KeyAtom(0)
    ReDim hCtrls(0)
    ReDim hAlts(0)
    ReDim hShifts(0)
End Sub
Private Function InHotValue(ByVal HotValue As Long, Optional ByVal hCtrl As Boolean = False, Optional ByVal hAlt As Boolean = False, Optional ByVal hShift As Boolean = False) As Long
    '判断热键值是否定义过
    '返回值:-1表示没有定义过,否则返回索引值
    Dim I As Long
   
    InHotValue = -1
   
    For I = 1 To UBound(HotKeys)
        If HotKeys(I) = HotValue Then       '先判断按键
            If hCtrl = hCtrls(I) And _
               hAlt = hAlts(I) And _
               hShift = hShifts(I) Then         '再判断功能键
                InHotValue = I
                Exit For
            End If
        End If
    Next I
End Function
Private Function GetTmpStr(ByVal NumValue As Long) As String
    '返回指定长度的随机字符串
    Dim I As Long
    Dim tmpStr As Long
   
    For I = 1 To NumValue
        Randomize
        tmpStr = Int(Rnd(1) * 70) + 60
        Select Case tmpStr
            Case 65 To 90, 97 To 122, 48 To 57
                If GetTmpStr = "" Then
                    GetTmpStr = Chr(tmpStr)
                Else
                    GetTmpStr = GetTmpStr & Chr(tmpStr)
                End If
        End Select
    Next I
End Function
Private Sub CtlTimer_Timer()
    '用于启动死循环的事件.在这里启动后,不会占用主界面的过程
    CtlTimer.Enabled = False
    If CtlTimer.Tag = "1" Then Exit Sub
    CtlTimer.Tag = "1"
    Call GetMessage
End Sub
Public Property Get frmHwnd() As Long
    frmHwnd = mfrmHwnd
End Property
Public Property Let frmHwnd(ByVal vNewValue As Long)
    mfrmHwnd = vNewValue
End Property

