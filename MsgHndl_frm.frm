VERSION 5.00
Begin VB.Form MsgHndl_frm 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "1000000000"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   1935
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3495
      Top             =   1200
   End
End
Attribute VB_Name = "MsgHndl_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Command1.Caption = "Start" Then
        Timer1.Tag = vbNullString
        Command1.Caption = "Stop"
        Text1.Text = vbNullString
        If IsNumeric(Text2.Text) Then
            Iterations = CLng(Text2.Text)
            Call RestartLoop(0)
        Else
            MsgBox "Enter how many test iterations you would like to perform in the topmost text box."
        End If
    Else
        Command1.Caption = "Start"
        blRunning = False
    End If
End Sub

Private Sub Form_Load()
    'MsgBox "Changing the message handling type while the" & vbNewLine & "loop is running will give inaccurate time results."
    blRunning = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blRunning = False
    'Call ShutDown
End Sub

Public Sub RestartLoop(Index As Integer)
    MsgType = Index
    I = 0
    
    'start the loop again
    blRunning = True
    Call Main
    If Timer1.Tag <> "Done" Then Timer1.Enabled = True
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call RestartLoop(MsgType + 1)
End Sub
