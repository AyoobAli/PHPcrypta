VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "FrmMain"
   ScaleHeight     =   3000
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label LblInfo 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblV 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0   Build: 0"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2700
      Width           =   1935
   End
   Begin VB.Label LblProgrammer 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3480
      MousePointer    =   10  'Up Arrow
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label LblSite 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label LblClose 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6480
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image ImgAbout 
      Height          =   3000
      Left            =   0
      Picture         =   "FrmAbout.frx":74F2
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
    Me.Width = ImgAbout.Width
    Me.Height = ImgAbout.Height
    LblV.Caption = App.Major & "." & App.Minor & "   Build: " & App.Revision
    
    LblInfo.Caption = "PHPcrypta is an open source software that allows you to encrypt a PHP code." & vbCrLf & vbCrLf _
    & "You are allowed to use or modify this software, but you are not allowed to sell it or any modified version of it."
End Sub

Private Sub ImgAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End If
End Sub

Private Sub LblClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub LblProgrammer_Click()
On Error Resume Next
    ShellExecute hwnd, "open", "http://www.Atyus.com", vbNullString, vbNullString, Empty
End Sub

Private Sub LblSite_Click()
On Error Resume Next
    ShellExecute hwnd, "open", "http://www.AXEsystems.com", vbNullString, vbNullString, Empty
End Sub
