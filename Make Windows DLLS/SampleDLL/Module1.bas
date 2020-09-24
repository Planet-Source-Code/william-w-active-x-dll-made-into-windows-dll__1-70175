Attribute VB_Name = "Module1"
Option Explicit
'Note if when adding a declare to another dll do not
'specify it in the .def file if you add it to the .def file it won't work
'the .def maker addin will do this for you
' Here Is Our Sample Test DLL
' If you want to present a typical ANSI function, taking parameters in as
'  ANSI and sending them back as ANSI, then it's necessary to
'  use StrConv(s, vbUnicode) on the incoming strings and StrConv(s, vbFromUnicode)
'  on the string values being returned.

Private Const WM_CLOSE = &H10
Public Const DLL_PROCESS_DETACH = 0
Public Const DLL_PROCESS_ATTACH = 1
Public Const DLL_THREAD_ATTACH = 2
Public Const DLL_THREAD_DETACH = 3

Public Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" ( _
      ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Long, _
      lparam As Any) As Long
      
Public Function DllMain(hInst As Long, fdwReason As Long, lpvReserved As Long) As Boolean
   Select Case fdwReason
      Case DLL_PROCESS_DETACH
         ' No per-process cleanup needed
      Case DLL_PROCESS_ATTACH
         DllMain = True
      Case DLL_THREAD_ATTACH
         ' No per-thread initialization needed
      Case DLL_THREAD_DETACH
         ' No per-thread cleanup needed
      Case Else
         DllMain = True ' dll should always return true unless it fails to load
   End Select
End Function

Public Function Increment(var As Integer) As Integer
   If Not IsNumeric(var) Then Err.Raise 5
   
   Increment = var + 1
End Function

Public Function Decrement(var As Integer) As Integer
   If Not IsNumeric(var) Then Err.Raise 5
   
   Decrement = var - 1
End Function
Public Function Test(ByVal hwnd As Long) As Long
   SendMessage hwnd, WM_CLOSE, 0, 0
End Function
Public Function Square(var As Long) As Long
   If Not IsNumeric(var) Then Err.Raise 5
   
   Square = var ^ 2
End Function

Public Function GetInfo(Str As String) As String
Str = StrConv(Str, vbUnicode)
Str = Str & " Sample DLL"
GetInfo = StrConv(Str, vbFromUnicode)
End Function

