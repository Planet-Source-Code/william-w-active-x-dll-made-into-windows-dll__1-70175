Attribute VB_Name = "Module1"
Option Explicit
Public SpecialLink As Boolean
Private intPos As Integer
Private fCPL As Boolean
Private fResource As Boolean
Private strCmd As String
Private strPath As String
Private strFileContents As String
Private strDefFile As String
Private strResFile As String
Private oFS As New Scripting.FileSystemObject
Private fld As Folder
Private fil As File
Private ts As TextStream
Private tsDef As TextStream

Public Sub DoLink()

   On Error Resume Next
  Dim ShellR As Long

   If SpecialLink Then
      ' Determine contents of .DEF file
      Set tsDef = oFS.OpenTextFile(strDefFile)

      strFileContents = tsDef.ReadAll

      If InStr(1, strFileContents, "CplApplet", vbTextCompare) > 0 Then
         fCPL = True
      End If

      ' Add module definition before /DLL switch
      intPos = InStr(1, strCmd, "/DLL", vbTextCompare)

      If intPos > 0 Then
         strCmd = Left(strCmd, intPos - 1) & " /DEF:" & Chr(34) & strDefFile & Chr(34) & " " & _
            Mid(strCmd, intPos)
      End If

      ' Include .RES file if one exists

      If fResource Then
         intPos = InStr(1, strCmd, "/ENTRY", vbTextCompare)
         strCmd = Left(strCmd, intPos - 1) & Chr(34) & strResFile & Chr(34) & " " & Mid(strCmd, _
            intPos)
      End If

      ' If Control Panel applet, change "DLL" extension to "CPL"

      If fCPL Then
         strCmd = Replace(strCmd, ".dll", ".cpl", 1, , vbTextCompare)
      End If

      ' Write linker options to output file
      ts.WriteLine "Command line arguments after modification:"
      ts.WriteBlankLines 1
      ts.WriteLine "   " & strCmd
      ts.WriteBlankLines 2
   End If

   ts.WriteLine "Calling LINK.EXE linker"
   Shell ("linklnk.exe " & strCmd)

   If Err.Number <> 0 Then
      ts.WriteLine "Error in calling linker..."
      Err.Clear
   End If

   ts.WriteBlankLines 1

   If InStr(1, strCmd, ".DLL" & Chr(34), vbTextCompare) Then
      If SpecialLink = True Then
         ts.WriteLine "Windows DLL File"
       Else
         ts.WriteLine "VB Active X DLL File"
      End If

   End If

   If InStr(1, strCmd, ".CPL" & Chr(34), vbTextCompare) Then
      If SpecialLink = True Then
         ts.WriteLine "Windows Control Panel Applet File"
       Else
         ts.WriteLine "Worthless Control Panel Applet File"
      End If

   End If
   If InStr(1, strCmd, ".EXE" & Chr(34), vbTextCompare) Then ts.WriteLine "Standard Windows EXE" & _
      " File"
   ts.WriteBlankLines 1
   ts.WriteLine "Returned from linker call"
   ts.Close
   Unload Form1

End Sub

Public Sub Main()

   On Error GoTo LinkErr
   strCmd = Command

   If strCmd = "" Then
      'Show Command Arguments Window
      Load Form2
      Exit Sub
   End If

   ' Determine if .DEF file exists
   '
   ' Extract path from first .obj argument
   intPos = InStr(1, strCmd, ".OBJ", vbTextCompare)
   strPath = Mid(strCmd, 2, intPos + 2)
   intPos = InStrRev(strPath, "\")
   strPath = Left(strPath, intPos - 1)

   Set ts = oFS.CreateTextFile(strPath & "\Lnklog.txt")
   'Start Log File
   ts.WriteLine "Beginning execution at " & Date & " " & Time()
   ts.WriteBlankLines 1

   ts.WriteLine "Command line arguments to LINK call:"
   ts.WriteBlankLines 1
   ts.WriteLine "   " & strCmd
   ts.WriteBlankLines 2

   ' Open folder
   Set fld = oFS.GetFolder(strPath)

   ' Get files in folder

   For Each fil In fld.Files

      If UCase(oFS.GetExtensionName(fil)) = "DEF" Then
         strDefFile = fil
         SpecialLink = True
      End If

      If UCase(oFS.GetExtensionName(fil)) = "RES" Then
         strResFile = fil
         fResource = True
      End If

      If SpecialLink And fResource Then Exit For
   Next

   ' Change command line arguments if flag set

   If SpecialLink = True Then
      If InStr(1, strCmd, ".DLL" & Chr(34), vbTextCompare) Or InStr(1, strCmd, ".CPL" & Chr(34), _
         vbTextCompare) Then
         Load Form1
       Else
         DoLink
      End If

    Else
      DoLink
   End If

   Exit Sub
LinkErr:
   MsgBox "Linker Error " & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Error"

End Sub

