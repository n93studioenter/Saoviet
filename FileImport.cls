VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFileImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: FileImport

Option Explicit

' Các thu?c tính
Private m_Name As String
Private m_Ngay As Date
Private m_SoHD As String
Private m_SoTK As String
Private m_Type As String
Private m_Path As String
Private m_NoTK As String
Private m_CoTK As String
Private m_ThueTK As String
Private m_DienGiai As String

' Thu?c tính Name
Public Property Get name() As String
    name = m_Name
End Property

Public Property Let name(Value As String)
    m_Name = Value
End Property

' Thu?c tính Ngay
Public Property Get ngay() As Date
    ngay = m_Ngay
End Property

Public Property Let ngay(Value As Date)
    m_Ngay = Value
End Property

' Thu?c tính SoHD
Public Property Get soHD() As String
    soHD = m_SoHD
End Property

Public Property Let soHD(Value As String)
    m_SoHD = Value
End Property

' Thu?c tính SoTK
Public Property Get soTK() As String
    soTK = m_SoTK
End Property

Public Property Let soTK(Value As String)
    m_SoTK = Value
End Property
Public Property Get patTH() As String
    patTH = m_Path
End Property

Public Property Let patTH(Value As String)
    m_Path = Value
End Property

' Thu?c tính Type
Public Property Get types() As String
types = m_Type
End Property

Public Property Let types(Value As String)
    m_Type = Value
End Property

Public Property Get notk() As String
    notk = m_NoTK
End Property
Public Property Let notk(Value As String)
    m_NoTK = Value
End Property
Public Property Get cotk() As String
    cotk = m_CoTK
End Property
Public Property Let cotk(Value As String)
    m_CoTK = Value
End Property

Public Property Get Diengiai() As String
    Diengiai = m_DienGiai
End Property
Public Property Let Diengiai(Value As String)
    m_DienGiai = Value
End Property
 
 Public Property Get ThueTK() As String
    ThueTK = m_ThueTK
End Property
Public Property Let ThueTK(Value As String)
    m_ThueTK = Value
End Property
