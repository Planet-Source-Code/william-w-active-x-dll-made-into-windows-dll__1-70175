VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export To .Def File 1.3"
   ClientHeight    =   2835
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6885
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Make Declare Helper File"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "Makes <Project name>_Declares.txt to help in implementing your functions"
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Export by Ordinal Only "
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      ToolTipText     =   "allows you to export by ordinal only and reduce the size of the export table in the resulting DLL"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2085
      ItemData        =   "frmAddIn.frx":08CA
      Left            =   120
      List            =   "frmAddIn.frx":08CC
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1185
      ItemData        =   "frmAddIn.frx":08CE
      Left            =   5520
      List            =   "frmAddIn.frx":08D0
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BGSOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NOTE: Only procedures in Modules can be exported into a DLL file! So, put all your DLL routines in a Module."
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select procedures to export to your .def file"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public AppPath As String
Private DecHelper() As String


Private Sub CancelButton_Click()

   Unload Me
   Connect.Hide

End Sub

Private Function FileEx(strx As String) As String

   'get only the file name from the complete path
   FileEx = Mid$(strx, Len(strx) - 3)

End Function

Private Function FileNm(strx As String) As String

   'get only the file name from the complete path
   FileNm = Mid$(strx, InStrRev(strx, "\") + 1, Len(strx))
   FileNm = Mid$(FileNm, 1, Len(FileNm) - 4)

End Function

Public Sub FLoad()

   On Error GoTo LdErr

  Dim objComponent As VBComponent
  Dim objMember As Member
  Dim a As Integer

   AppPath = VBInstance.ActiveVBProject.FileName

   If AppPath = "" Then
      MsgBox "Please save your project before choosing what you want to export.", vbInformation, "Project Data Not Found"
      Unload Me
      Connect.Hide
      Exit Sub
   End If

   a = 0
   List1.Clear
   List2.Clear

   AppPath = Left$(AppPath, Len(AppPath) - 4) & ".def"

   For Each objComponent In VBInstance.ActiveVBProject.VBComponents

      If objComponent.Type = vbext_ct_StdModule Then

         For Each objMember In objComponent.CodeModule.Members

            If objMember.Type = vbext_mt_Method Then

               If IsDeclare(objMember.Name, objComponent) = False Then  ' We don't want Declares added to the .def file

                  List2.AddItem objMember.Name & " (defined in: " & objComponent.Name & ", Scope: " & Getscope(objMember.Scope) & " )"
                  'MsgBox getdeclare(objMember.Name, objComponent, objMember.Scope)
                  ReDim Preserve DecHelper(a)
                  'get each procedures inputs/outputs and store them to a string array for later
                  DecHelper(a) = getdeclare(objMember.Name, objComponent, objMember.Scope)
                  a = a + 1

                  If (objMember.Scope = vbext_Friend Or objMember.Scope = vbext_Public) Then

                     List2.Selected(List2.ListCount - 1) = True

                  End If
                  'Private members won't be checked automatically though I don't know why you would need to
                  'export a private member I'll still let you do so if desired
               End If

            End If

         Next

      End If
   Next

   For a = List2.ListCount - 1 To 0 Step -1
      'check the box next to each procedure.
      List1.AddItem List2.List(a)
      List1.Selected(List1.ListCount - 1) = List2.Selected(a)

   Next
   Exit Sub
LdErr:
   MsgBox "Load Error " & Err.Description & " " & Err.Number, vbCritical Or vbOKOnly, "Make .Def File Error"
   Unload Me
   Connect.Hide

End Sub

Private Function getdeclare(Name As String, objComponent As VBComponent, Scope As Long) As String
'Probably could make this shorter but i want to be sure not to pick up the wrong
'functions/subs
   
   ' get functions or subs input and output params
  Dim stline As Long

   stline = 0

   If objComponent.CodeModule.Find(Getscope(Scope) & " Function " & Name & "(", stline, 0, 0, 0, False, True, False) = True Then
      'if you use a function and don't give it a type ie. Function(a as long) as long
      'then this won't find it properly
      getdeclare = objComponent.CodeModule.Lines(stline, 10)
      getdeclare = Mid(getdeclare, 1, InStr(InStr(2, getdeclare, ") As ", vbTextCompare) + 5, getdeclare, " ", vbTextCompare))
      getdeclare = Mid(getdeclare, 1, InStrRev(getdeclare, vbCrLf))

      If getdeclare = "" Then
         'but if the above doesn't find it properly this will find a function without a return type
         getdeclare = objComponent.CodeModule.Lines(stline, 10)
         getdeclare = Mid(getdeclare, 1, InStr(2, getdeclare, ")", vbTextCompare))
      End If

    ElseIf objComponent.CodeModule.Find(Getscope(Scope) & " Sub " & Name & "(", stline, 0, 0, 0, False, True, False) = True Then
      getdeclare = objComponent.CodeModule.Lines(stline, 10)
      getdeclare = Mid(getdeclare, 1, InStr(2, getdeclare, ")", vbTextCompare))
      'if this is a sub we'll find the last ')' and end it there
      'below is the same as above but it finds the subs/functions without a typed scope
      'Ie: Above finds Public Function Test(a as long)  but below finds Function Test(a as long)
    ElseIf objComponent.CodeModule.Find("Function " & Name & "(", stline, 0, 0, 0, False, True, False) = True Then
      'if you use a function and don't give it a type ie. Function(a as long) as long
      'then this won't find it properly
      getdeclare = objComponent.CodeModule.Lines(stline, 10)
      getdeclare = Mid(getdeclare, 1, InStr(InStr(2, getdeclare, ") As ", vbTextCompare) + 5, getdeclare, " ", vbTextCompare))
      getdeclare = Mid(getdeclare, 1, InStrRev(getdeclare, vbCrLf))

      If getdeclare = "" Then
         'but if the above doesn't find it properly this will find a function without a return type
         getdeclare = objComponent.CodeModule.Lines(stline, 10)
         getdeclare = Mid(getdeclare, 1, InStr(2, getdeclare, ")", vbTextCompare))
      End If

    ElseIf objComponent.CodeModule.Find("Sub " & Name & "(", stline, 0, 0, 0, False, True, False) = True Then
      getdeclare = objComponent.CodeModule.Lines(stline, 10)
      getdeclare = Mid(getdeclare, 1, InStr(2, getdeclare, ")", vbTextCompare))
      'if this is a sub we'll find the last ')' and end it there
   End If

   If getdeclare <> "" Then getdeclare = MakeDeclare(getdeclare)

End Function

Private Function Getscope(Scope As Long) As String

   ' obviously it outputs the scope as a string

   Select Case Scope
    Case 1:
      Getscope = "Private"

    Case 2:
      Getscope = "Public"

    Case 3:
      Getscope = "Friend"

    Case Else:
      Getscope = "Unknown"
   End Select

End Function

Private Function IsDeclare(Name As String, objComponent As VBComponent) As Boolean

   '^ checks for declared functions/subs, you can't add declared functions/subs to the .def file otherwise the dll will fail

   If objComponent.CodeModule.Find("Declare Function " & Name & " Lib", 0, 0, 0, 0, True, True, False) = False And objComponent.CodeModule.Find("Declare Sub " & Name & "Lib", 0, 0, 0, 0, True, True, False) = False Then
      IsDeclare = False
    Else
      IsDeclare = True
   End If

End Function

Private Function MakeDeclare(Inpt As String) As String

   On Error GoTo ExitFun
'adds Declare & Lib "<filename>"
  Dim dtype As Long
  Dim S1 As Long
  Dim S2 As Long

   dtype = 0

   If InStr(1, Mid(Inpt, 1, 20), "Function ", vbTextCompare) Then dtype = 1
   If InStr(1, Mid(Inpt, 1, 15), "Sub ", vbTextCompare) Then dtype = 2

   Select Case dtype

    Case 1:
      S1 = InStr(1, Mid(Inpt, 1, 20), "Function ", vbTextCompare)
      S2 = InStr(S1, Inpt, "(", vbTextCompare)

      MakeDeclare = Mid(Inpt, 1, S1 - 1) & "Declare " & Mid(Inpt, S1, S2 - S1) & " Lib " & Chr(34) & FileNm(VBInstance.ActiveVBProject.BuildFileName) & FileEx(VBInstance.ActiveVBProject.BuildFileName) & Chr(34) & " " & Mid(Inpt, S2)

    Case 2:
      S1 = InStr(1, Mid(Inpt, 1, 20), "Sub ", vbTextCompare)
      S2 = InStr(S1, Inpt, "(", vbTextCompare)

      MakeDeclare = Mid(Inpt, 1, S1 - 1) & "Declare " & Mid(Inpt, S1, S2 - S1) & " Lib " & Chr(34) & FileNm(VBInstance.ActiveVBProject.BuildFileName) & FileEx(VBInstance.ActiveVBProject.BuildFileName) & Chr(34) & " " & Mid(Inpt, S2)

    Case Else:
      MakeDeclare = Inpt

   End Select

   Exit Function

ExitFun:
   MakeDeclare = Inpt

End Function

Private Sub OKButton_Click()

   On Error GoTo SaveError
  Dim a As Long

  Dim strTemp
  Dim OptionalCmd As String

  Dim objComponent As VBComponent
  Dim objMember As Member

   'open the .def file for the project - this outputs all
   'the exports you select in the end dll file.
   Open AppPath For Output As #2
   Print #2, "Name " & FileNm(VBInstance.ActiveVBProject.BuildFileName)
   Print #2, "LIBRARY " & VBInstance.ActiveVBProject.Name
   Print #2, "EXPORTS"

   'go throgh all procedures in the list box. If it is
   'checked, write the name of it into the file along with '@' and the ordinal Position
   ' if ordinal only is selected then The optional NONAME keyword allows you to export
   'by ordinal only and reduce the size of the export table in the resulting DLL.
   'However, if you want to use GetProcAddress on the DLL, you must know the ordinal
   'because the name will not be valid.
   'See: http://msdn2.microsoft.com/en-us/library/hyx1zcd3(vs.71).aspx

   If Check1.Value = 1 Then
      OptionalCmd = " NONAME"
    Else
      OptionalCmd = ""
   End If

   For a = 0 To List1.ListCount - 1

      If List1.Selected(a) = True Then
         Print #2, "    " & Split(List1.List(a), " ")(0) & " @" & a + 1 & OptionalCmd
      End If

   Next

   If Check2.Value = 1 Then 'Make Declare Helper File

      AppPath = VBInstance.ActiveVBProject.FileName
      AppPath = Left$(AppPath, Len(AppPath) - 4) & "_declares.txt"
      Open AppPath For Output As #1
      Print #1, "Declare Helper File"
      Print #1, "Name " & FileNm(VBInstance.ActiveVBProject.BuildFileName)
      Print #1, "LIBRARY " & VBInstance.ActiveVBProject.Name
      Print #1, ""
      Print #1, ""

      For a = 0 To List1.ListCount - 1

         If List1.Selected(a) = True Then
            Print #1, DecHelper(List1.ListCount - 1 - a)
            Print #1, ""
         End If

      Next

   End If

   Close 'close open files

   MsgBox AppPath & vbCrLf & " Successfully Created", vbOKOnly, "Make .Def Success"
   Unload Me

   Connect.Hide
   Exit Sub
SaveError:
   Close
   MsgBox "An error occured while writing the definition file: " & Err.Description & " (" & Err.Number & ")", vbOKOnly, "Error"
   Unload Me
   Connect.Hide

End Sub

