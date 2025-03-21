VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Khai b�o c�c bi?n d? luu th�ng tin v? file
Private pFileName As String
Private pFilePath As String
Private pCreationDate As Date

' Thu?c t�nh cho T�n File
Public Property Let FileName(ByVal FileName As String)
    pFileName = FileName
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property

' Thu?c t�nh cho �u?ng D?n File
Public Property Let FilePath(ByVal FilePath As String)
    pFilePath = FilePath
End Property

Public Property Get FilePath() As String
    FilePath = pFilePath
End Property

' Thu?c t�nh cho Ng�y T?o
Public Property Let CreationDate(ByVal CreationDate As Date)
    pCreationDate = CreationDate
End Property

Public Property Get CreationDate() As Date
    CreationDate = pCreationDate
End Property
