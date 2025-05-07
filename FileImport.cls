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
Private m_ID As String
Private m_Name As String
Private m_MST As String
Private m_Ngay As Date
Private m_SoHD As String
Private m_KHHD As String
Private m_SoTK As String
Private m_Type As String
Private m_Path As String
Private m_NoTK As String
Private m_CoTK As String
Private m_ThueTK As String
Private m_DienGiai As String
Private m_TongTien As String
Private m_TgTCThue As String
Private m_TgTThue As String
Private m_Vat As String
Private m_SohieuTP As String
Public Property Get id() As String
    id = m_ID
End Property

Public Property Let id(Value As String)
    m_ID = Value
End Property
' Thu?c tính Name
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property
' Thu?c tính Name
Public Property Get mst() As String
    mst = m_MST
End Property

Public Property Let mst(Value As String)
    m_MST = Value
End Property

' Thu?c tính Ngay
Public Property Get ngay() As Date
    ngay = m_Ngay
End Property

Public Property Let ngay(Value As Date)
    m_Ngay = Value
End Property

' Thu?c tính SoHD
Public Property Get sohd() As String
    sohd = m_SoHD
End Property

Public Property Let sohd(Value As String)
    m_SoHD = Value
End Property
Public Property Get khHD() As String
    khHD = m_KHHD
End Property

Public Property Let khHD(Value As String)
    m_KHHD = Value
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

Public Property Get diengiai() As String
    diengiai = m_DienGiai
End Property
Public Property Let diengiai(Value As String)
    m_DienGiai = Value
End Property
 
 Public Property Get ThueTK() As String
    ThueTK = m_ThueTK
End Property
Public Property Let ThueTK(Value As String)
    m_ThueTK = Value
End Property
 Public Property Get TongTien() As String
    TongTien = m_TongTien
End Property
Public Property Let TongTien(Value As String)
    m_TongTien = Value
End Property
 Public Property Get TgTCThue() As String
    TgTCThue = m_TgTCThue
End Property
Public Property Let TgTCThue(Value As String)
    m_TgTCThue = Value
End Property
Public Property Get TgTThue() As String
    TgTThue = m_TgTThue
End Property
Public Property Let TgTThue(Value As String)
    m_TgTThue = Value
End Property
 Public Property Get VAT() As String
    VAT = m_Vat
End Property
Public Property Let VAT(Value As String)
    m_Vat = Value
End Property
 Public Property Get sohieutp() As String
    sohieutp = m_SohieuTP
End Property
Public Property Let sohieutp(Value As String)
    m_SohieuTP = Value
End Property
