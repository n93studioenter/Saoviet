VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Le Duc Hong         http://www.vovisoft.com
Option Explicit
Private mDOMTextFile As DOMDocument  ' Document Object Model
Private mXMLPath As String  ' XML filename
Public Function ReadUnicode(TXMLPath)
' Read Unicode text from XML file
  Dim objTextFileRoot As IXMLDOMElement
  Set mDOMTextFile = New DOMDocument
  ' Remember the XML file name to update later
  mXMLPath = TXMLPath
  ' Read the XML file and create a DOM
  mDOMTextFile.Load mXMLPath
  'start at the root element of the XML
   Set objTextFileRoot = mDOMTextFile.documentElement
   ' Return the root node's text
   ReadUnicode = objTextFileRoot.nodeTypedValue
End Function

Public Sub WriteUnicode(OutText, Optional TXMLPath)
' Update Unicode Text of same XML file or write to another XML file
    If IsMissing(TXMLPath) Then
        ' Assign the Text to write out to the Root node of DOM
          mDOMTextFile.documentElement.Text = OutText
        ' Update the XML file
         mDOMTextFile.Save mXMLPath
     Else
        ' Read a dummy file to create a DOM
        ReadUnicode GetLocalDirectory & "Dummy.xml"
        ' Assign the Text to write out to the Root node of DOM
          mDOMTextFile.documentElement.Text = OutText
        ' Write the Text to the given XML file
         mDOMTextFile.Save TXMLPath
     End If
End Sub


