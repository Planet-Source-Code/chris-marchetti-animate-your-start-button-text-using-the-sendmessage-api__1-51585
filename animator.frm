VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Animator"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   1320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Word5"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Word4"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Word3"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Word2"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Word1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
    
    StartBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
    
sCaption = Text1
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
    
    StartBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
    
sCaption = Text2
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
    
    StartBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
    
sCaption = Text3
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
    
    StartBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
    
sCaption = Text4
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Dim StartBar As Long
Dim StartBarText As Long
Dim sCaption As String
    
    StartBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartBarText = FindWindowEx(StartBar, 0&, "button", vbNullString)
    
sCaption = Text5
SendMessageSTRING StartBarText, WM_SETTEXT, 256, sCaption
Timer1.Enabled = True
Timer5.Enabled = False
End Sub
