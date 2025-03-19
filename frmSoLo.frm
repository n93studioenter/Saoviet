VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSoLo 
   Caption         =   "Sè l« & H¹n dïng"
   ClientHeight    =   780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3840
   LinkTopic       =   "Form3"
   ScaleHeight     =   780
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsolo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "..."
      Top             =   240
      Width           =   1815
   End
   Begin MSMask.MaskEdBox txtngaynhap 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmSoLo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtsolo.SelStart = 0
txtsolo.SelLength = Len(txtsolo.Text)
txtngaynhap.SelStart = 0
txtngaynhap.SelLength = Len(txtngaynhap.Text)
End Sub

Public Sub Form_Unload(Cancel As Integer)
Unload frmSoLo
End Sub

Private Sub txtngaynhap_KeyPress(KeyAscii As Integer)
If txtsolo.Text = "" Then
MsgBox "Ban phai nhap so lo"
Else
 If KeyAscii = 13 Then Me.Visible = False ' Me.Hide
End If

End Sub

Private Sub txtngaynhap_LostFocus()
txtngaynhap.SelStart = 0
txtngaynhap.SelLength = Len(txtngaynhap.Text)

End Sub

Private Sub txtsolo_LostFocus()
 txtsolo.SelStart = 0
txtsolo.SelLength = Len(txtsolo.Text)
End Sub
