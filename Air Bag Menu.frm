VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Air Bag Menu.frx":0000
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3645
      Index           =   1
      Left            =   8520
      Picture         =   "Air Bag Menu.frx":DCB5A
      ScaleHeight     =   3645
      ScaleWidth      =   3600
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.PictureBox Mascara 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8370
      Left            =   7920
      Picture         =   "Air Bag Menu.frx":10770C
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   8100
   End
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3825
      Index           =   5
      Left            =   8640
      Picture         =   "Air Bag Menu.frx":1E4266
      ScaleHeight     =   3825
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3435
      Index           =   4
      Left            =   8640
      Picture         =   "Air Bag Menu.frx":211FC8
      ScaleHeight     =   3435
      ScaleWidth      =   3870
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   3870
   End
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3480
      Index           =   3
      Left            =   8520
      Picture         =   "Air Bag Menu.frx":23D632
      ScaleHeight     =   3480
      ScaleWidth      =   3765
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4095
      Index           =   2
      Left            =   8640
      Picture         =   "Air Bag Menu.frx":268394
      ScaleHeight     =   4095
      ScaleWidth      =   3825
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.PictureBox Seleção 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3570
      Index           =   0
      Left            =   8520
      Picture         =   "Air Bag Menu.frx":29B6D6
      ScaleHeight     =   3570
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press ESC to exit"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use SHIFT and Drag to move the form"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION As Integer = 2
Private Const WM_NCLBUTTONDOWN As Integer = &HA1
Private Const LWA_COLORKEY As Integer = &H1
Private Const LWA_ALPHA As Integer = &H2
Private Const GWL_EXSTYLE As Integer = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Cor(0 To 5) As Long
Private ultima_cor As Long
Private ultima As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()

  Dim Ret As Long
  Dim Col As Long
    
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
  
    Col = RGB(0, 0, 0)

    SetLayeredWindowAttributes Me.hwnd, Col, 50, LWA_COLORKEY

    Cor(0) = Mascara.Point(317, 216)
    Cor(1) = Mascara.Point(352, 303)
    Cor(2) = Mascara.Point(317, 375)
    Cor(3) = Mascara.Point(186, 381)
    Cor(4) = Mascara.Point(169, 318)
    Cor(5) = Mascara.Point(219, 229)
    
    Seleção(0).Move 100, 57
    Seleção(1).Move 219, 55
    Seleção(2).Move 272, 137
    Seleção(3).Move 214, 287
    Seleção(4).Move 83, 287
    Seleção(5).Move 38, 154

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim ReturnVal As Long
  Dim r As Integer
    If Shift Then
        X = ReleaseCapture()
        ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        Exit Sub
    End If
    Seleção(ultima).Visible = False

    Select Case ultima
      Case 0
        r = MsgBox("Option 6", vbInformation, "")
      Case 1
        r = MsgBox("Option 1", vbInformation, "")
      Case 2
        r = MsgBox("Option 2", vbInformation, "")
      Case 3
        r = MsgBox("Option 3", vbInformation, "")
      Case 4
        r = MsgBox("Option 4", vbInformation, "")
      Case 5
        r = MsgBox("Option 5", vbInformation, "")
    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Mascara.Point(X, Y) <> ultima_cor Then
        ultima_cor = Mascara.Point(X, Y)
        Seleção(ultima).Visible = False
        Select Case Mascara.Point(X, Y)
          Case Cor(0)
            ultima = 1
            Seleção(ultima).Visible = True
          Case Cor(1)
            ultima = 2
            Seleção(ultima).Visible = True
          Case Cor(2)
            ultima = 3
            Seleção(ultima).Visible = True
          Case Cor(3)
            ultima = 4
            Seleção(ultima).Visible = True
          Case Cor(4)
            ultima = 5
            Seleção(ultima).Visible = True
          Case Cor(5)
            ultima = 0
            Seleção(ultima).Visible = True
          Case -1
            Seleção(ultima).Visible = False
        End Select
    End If

End Sub



