Attribute VB_Name = "MsgHndl_mdl"
Option Explicit

'Message handling
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

'For timing
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'Constants
'- PostMessage
Private Const PM_NOREMOVE = &H0
Private Const PM_REMOVE = &H1

'- Window Messages
Private Const WM_QUIT = &H12

'- GetQueueStatus
Private Const QS_HOTKEY = &H80
Private Const QS_KEY = &H1
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSEMOVE = &H2
Private Const QS_PAINT = &H20
Private Const QS_POSTMESSAGE = &H8
Private Const QS_SENDMESSAGE = &H40
Private Const QS_TIMER = &H10
Private Const QS_ALLPOSTMESSAGE = &H100
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

'Types, Enums
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MSG
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    point As POINTAPI
End Type

Public Enum MessageHandleType
    GetMsg = 0
    PeekMsg = 1
    DoEvnts = 2
    QStatus = 3
End Enum

'Globals
Public blRunning As Boolean
Public MsgType As MessageHandleType
Public Iterations As Long

'the loop counter
Public I As Long

Public Sub Main()
    Dim RetVal As Long
    Dim Freq As Currency
    Dim StartTime As Currency
    Dim StopTime As Currency
    Dim RunTime As Currency
    Dim Message As MSG
    Dim sMsgPmpType As String
    
    'Debug.Print "Entered Main()"
    
    Select Case MsgType
        Case 0:
            sMsgPmpType = "GetQueueStatus"
        Case 1:
            sMsgPmpType = "PeekMessage"
        Case 2:
            sMsgPmpType = "DoEvents"
        Case 3:
            sMsgPmpType = "GetMessage"
        Case Else:
            sMsgPmpType = "Done"
            blRunning = False
            MsgHndl_frm.Timer1.Tag = "Done"
    End Select
    
    MsgHndl_frm.Caption = sMsgPmpType
    
    QueryPerformanceFrequency Freq
    
    QueryPerformanceCounter StartTime
    Do While blRunning
        Select Case MsgType
            Case 0:     'GetQueueStatus
                RetVal = GetQueueStatus(QS_ALLEVENTS Or QS_ALLINPUT)
                'retval will now hold the message number
                RetVal = HiWord(RetVal)
            Case 1:     'PeekMessage
                '*** use the PeekMessage version if you want to use 100% CPU
                '*** frex to do background processing that doesn't rely on
                '*** windows messages
                If PeekMessage(Message, 0&, 0&, 0&, PM_REMOVE) Then
                    Call TranslateMessage(Message)
                    Call DispatchMessage(Message)
                End If
            Case 2:     'DoEvents
                '*** the 'pure vb' way to do this
                DoEvents
            Case 3:     'GetMessage
                '*** use the GetMessage version if you only want to do
                '*** processing if there's a message
                If GetMessage(Message, 0&, 0&, 0&) Then
                    Call TranslateMessage(Message)
                    Call DispatchMessage(Message)
                End If
            Case Else:
                I = 0
                blRunning = False
                Exit Do
        End Select
        
        'code body -- this is where you put all your code
        'frex: calls to render, to get KBState, DBUpdate, etc
        I = I + 1
        
        If I = Iterations Then
            'exit the loop
            blRunning = False
        End If
    Loop

    'go on to the next MessagePump type
    'If I <> 0 And blRunning = False Then Call MsgHndl_frm.Option1_Click(MsgType + 1)
    
    QueryPerformanceCounter StopTime
    
    'Debug.Print "Exited the loop."
    If I = 0 Or StartTime = 0 Then Exit Sub
    
    RunTime = (StopTime - StartTime) / Freq
    If RunTime <> 0 Then
        'Debug.Print StopTime, StartTime, Freq
        MsgHndl_frm.Text1.Text = MsgHndl_frm.Text1.Text & sMsgPmpType & vbNewLine
        MsgHndl_frm.Text1.Text = MsgHndl_frm.Text1.Text & "Runtime in seconds: " & RunTime & ", Loop iterations: " & I & vbNewLine
        MsgHndl_frm.Text1.Text = MsgHndl_frm.Text1.Text & "Iterations per second: " & Format$(I / RunTime, "#,##0.0#####") & vbNewLine
        MsgHndl_frm.Text1.Text = MsgHndl_frm.Text1.Text & vbNewLine
        'Debug.Print sMsgPmpType
        'Debug.Print "Runtime in seconds: " & RunTime & ", Loop iterations: " & I
        'Debug.Print "Iterations per second: " & I / RunTime
    End If
    
    'we fell out of the loop, clean everything up
    'Call ShutDown
End Sub

Public Sub ShutDown()
    Dim Form As Form
    
    'destructor code that holds a form reference goes here
    
    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next
    
    'any other destructor code goes here
    
    'This is not normally necessary, we're calling a form destructor from a form though
    'which will just reactivate the form
    End
End Sub

Public Function LoWord(ByVal LongVal As Long) As Integer
    LoWord = LongVal And &HFFFF&
End Function

Public Function HiWord(ByVal LongVal As Long) As Integer
    If LongVal <> 0 Then HiWord = LongVal \ &H10000 And &HFFFF&
End Function
