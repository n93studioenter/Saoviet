Attribute VB_Name = "Module2"
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Public Function ReadTextFile(FileName) As String
'   Dim Fs As FileSystemObject
'   Dim TS As TextStream
'   '  Create a FileSystem Object
'   Set Fs = CreateObject("Scripting.FileSystemObject")
'   ' Open TextStream for Input
'   Set TS = Fs.OpenTextFile(FileName, ForReading, False, TristateUseDefault)
'   ReadTextFile = TS.ReadAll  ' Read the whole content of the text file in one stroke
'   TS.Close ' Close the Text Stream
'   Set Fs = Nothing  ' Dispose FileSystem Object
End Function
Public Sub WriteTextFile(FileName, StrOutText)
'   Dim Fs As FileSystemObject
'   Dim TS As TextStream
'   '  Create a FileSystem Object
'   Set Fs = CreateObject("Scripting.FileSystemObject")
'   ' Open TextStream for Output, create file if necesssary
'   Set TS = Fs.OpenTextFile(FileName, ForWriting, True, TristateUseDefault)
'   TS.Write StrOutText  ' Write the whole StrOutText string in one stroke
'   TS.Close ' Close the Text Stream
'   Set Fs = Nothing  ' Dispose FileSystem Object
End Sub
Function GetLocalDirectory() As String
' Obtain the folder where the program resides
Dim TStr
  ' Get folder where the Exe of this program resides
  TStr = App.path
  ' Append a backslash if folder does not end with one
  If Right(TStr, 1) <> "\" Then TStr = TStr & "\"
  GetLocalDirectory = TStr ' Return the folder
End Function
Sub PopulateListBoxFromFile(ByVal AListBox, ByVal AFileName As String, ByVal HasFolder As Boolean)
' Display a text file in a listbox
   Dim theFileName As String
   Dim FileNum, aLine
   ' If Filename is not fullpath then use LocalDirectory
   If HasFolder Then
      theFileName = AFileName
   Else
      theFileName = GetLocalDirectory & AFileName
   End If
   ' Get out if the input file does not exist
   If Dir(theFileName) = "" Then Exit Sub
   FileNum = FreeFile
   AListBox.Clear  ' Clear the listbox
   ' Open input file
   Open theFileName For Input As FileNum
   ' Read till End-Of-File
   Do While Not EOF(FileNum)
      Line Input #FileNum, aLine  '  Read a Text line
      AListBox.AddItem aLine      ' Add the line to the listbox
   Loop
   Close FileNum  ' Close the input file
End Sub
Sub SaveListBoxToFile(ByVal AListBox, ByVal AFileName As String, ByVal HasFolder As Boolean)
' Save content of a listbox to a text file
   Dim theFileName As String
   Dim FileNum, i
   ' Get out if there's nothing to save
   If AListBox.ListCount = 0 Then Exit Sub
   ' If Filename is not fullpath then use LocalDirectory
   If HasFolder Then
      theFileName = AFileName
   Else
      theFileName = GetLocalDirectory & AFileName
   End If
   FileNum = FreeFile
   ' Open output file
   Open theFileName For Output As FileNum
   ' Write each line of the listbox to the file
   For i = 0 To AListBox.ListCount - 1
      Print #FileNum, AListBox.List(i)
   Next
   Close FileNum  ' Close the output file
End Sub



