VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close from DLL"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSquare 
      Caption         =   "Square"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncrement 
      Caption         =   "Increment"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecrement 
      Caption         =   "Decrement"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Increment Lib "MathLib.dll" (var As Integer) As Integer

Private Declare Function Decrement Lib "MathLib.dll" (var As Integer) As Integer

Private Declare Function Square Lib "MathLib.dll" (var As Long) As Long

Private Declare Function Test Lib "MathLib.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetInfo Lib "MathLib.dll" (Str As String) As String

Dim incr As Integer
Dim decr As Integer
Dim sqr As Long

Private Sub cmdDecrement_Click()
   decr = Decrement(decr)
   cmdDecrement.Caption = "x = " & CStr(decr)
End Sub

Private Sub cmdIncrement_Click()
   incr = Increment(incr)
   cmdIncrement.Caption = "x = " & CStr(incr)
End Sub

Private Sub cmdSquare_Click()
   sqr = Square(sqr)
   cmdSquare.Caption = "x = " & CStr(sqr)
End Sub

Private Sub Command1_Click()
MsgBox GetInfo(Date$)
End Sub

Private Sub Command2_Click()
Test Me.hwnd
End Sub

Private Sub Form_Load()
   incr = 1
   decr = 100
   sqr = 2
End Sub

