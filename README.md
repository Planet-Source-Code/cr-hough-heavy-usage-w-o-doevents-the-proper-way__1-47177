<div align="center">

## Heavy usage w/o DoEvents \(the proper way\)


</div>

### Description

Never use DoEvents again. This is a standard C/C++ Message Pump. Compare each of these three methods against each other (including DoEvents) and watch your CPU usage. The code is pretty self-explanatory I think.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-07-29 02:15:02
**By**             |[CR Hough](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cr-hough.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Heavy\_usag1621187292003\.zip](https://github.com/Planet-Source-Code/cr-hough-heavy-usage-w-o-doevents-the-proper-way__1-47177/archive/master.zip)





### Source Code

  Put this in the General Declarations area for your module:<br>
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin
As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long<br>
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As
Long, ByVal wMsgFilterMax As Long) As Long<br>
Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long<br>
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long<br>
<br>
Private Const PM_NOREMOVE = &H0<br>
Private Const PM_REMOVE = &H1<br>
Private Const WM_QUIT = &H12<br>
<br>
Private Type POINTAPI<br>
    X As Long<br>
    Y As Long<br>
End Type<br>
<br>
Private Type MSG<br>
    hwnd As Long<br>
    Message As Long<br>
    wParam As Long<br>
    lParam As Long<br>
    time As Long<br>
    point As POINTAPI<br>
End Type<br>
<br>
<br>
  Put the following in your sub main (after setting blRunning = True):<br>
<pre>
'***********************************************************
'*** Try commenting and uncommenting each of these three ***
'*** methods. Watch your CPU usage when you run each of ***
'*** them. Then go back to using GetMessage.   ***
'***********************************************************
</pre>
Do While blRunning<br>
    '*** use the PeekMessage version if you want to use 100% CPU<br>
    '*** frex to do background processing that doesn't rely on<br>
    '*** windows messages<br>
    'If PeekMessage(Message, 0&, 0&, 0&, PM_REMOVE) Then<br>
    '    Call TranslateMessage(Message)<br>
    '    Call DispatchMessage(Message)<br>
    'End If<br>
<br>
    '*** use the GetMessage version if you only want to do<br>
    '*** processing if there's a message<br>
    If GetMessage(Message, 0&, 0&, 0&) Then<br>
    Call TranslateMessage(Message)<br>
    Call DispatchMessage(Message)<br>
    End If<br>
<br>
    '*** the 'pure vb' poor way to do this<br>
    'DoEvents<br>
Loop<br>

