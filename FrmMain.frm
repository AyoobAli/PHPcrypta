VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PHPcrypta"
   ClientHeight    =   8055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13095
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "FrmMain"
   ScaleHeight     =   8055
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEncrypt 
      Caption         =   "Encrypt"
      Enabled         =   0   'False
      Height          =   280
      Left            =   11640
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C7B3A7&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C7B3A7&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox TxtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C7B3A7&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar SttBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7800
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15822
            Text            =   "Select File..."
            TextSave        =   "Select File..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "22/05/2013"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:39 PM"
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton OptLvl 
      BackColor       =   &H00C7B3A7&
      Caption         =   "Advanced"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton OptLvl 
      BackColor       =   &H00C7B3A7&
      Caption         =   "Password"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton OptLvl 
      BackColor       =   &H00C7B3A7&
      Caption         =   "Simple"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CDlog 
      Left            =   6360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse..."
      Height          =   280
      Left            =   11880
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00C7B3A7&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   6495
   End
   Begin VB.TextBox TxtCode 
      BackColor       =   &H00C7B3A7&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   12855
   End
   Begin VB.Label LblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   270
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   270
      Width           =   375
   End
   Begin VB.Shape ShpBG 
      BackColor       =   &H00C7B3A7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AD7B5E&
      Height          =   495
      Index           =   1
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8175
   End
   Begin VB.Shape ShpBG 
      BackColor       =   &H00C7B3A7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AD7B5E&
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4575
   End
   Begin VB.Shape ShpBG 
      BackColor       =   &H00D6CDC7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AD7B5E&
      Height          =   615
      Index           =   3
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   7695
   End
   Begin VB.Shape ShpBG 
      BackColor       =   &H00C7B3A7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AD7B5E&
      Height          =   6855
      Index           =   4
      Left            =   -120
      Top             =   1200
      Width           =   13335
   End
   Begin VB.Shape ShpBG 
      BackColor       =   &H00D6CDC7&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00AD7B5E&
      Height          =   615
      Index           =   2
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4095
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "&File"
      Begin VB.Menu Mnu_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Mnu_Open 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu Mnu_nl1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Mnu_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu Mnu_Enc 
         Caption         =   "Encryption Level"
         Begin VB.Menu Mnu_lvl 
            Caption         =   "Simple"
            Checked         =   -1  'True
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu Mnu_lvl 
            Caption         =   "Password"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu Mnu_lvl 
            Caption         =   "Advanced"
            Index           =   2
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu Mnu_nl2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Start 
         Caption         =   "Start Encrypting"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Mnu_Help 
      Caption         =   "Help"
      Begin VB.Menu Mnu_Instructions 
         Caption         =   "Instructions"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu Mnu_nl3 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_About 
         Caption         =   "About"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBrowse_Click()
On Error GoTo Err:

Dim FFnum As Integer
Dim ConfirmMSG

FFnum = FreeFile

If Right(SttBar.Panels(1).Text, 1) = "*" Then
    ConfirmMSG = MsgBox("Are you sure you want to open a new file ?" & vbCrLf & "All Changes will be lost.", vbQuestion + vbYesNo, "Open New File")
    
    If ConfirmMSG = vbNo Then GoTo Err:
End If


With CDlog
    .FileName = ""
    .DialogTitle = "Select PHP File"
    .Filter = "(.php, .php3) PHP Files|*.php;.php3|(.txt) Text Files|*.txt|(*.*) All Files|*.*"
    
    .CancelError = True
    .ShowOpen

    TxtFile.Text = .FileName
    
    Open TxtFile.Text For Input As #FFnum
    strTheData = StrConv(InputB(LOF(FFnum), 1), vbUnicode)
    Close #FFnum
    
    TxtCode.Text = strTheData
    
    If InStr(TxtCode.Text, "<?") > 0 Or InStr(TxtCode.Text, "?>") > 0 Then
        wMsg = MsgBox("Do you want to delete all '<?php' and '?>' tags ?", vbYesNo + vbQuestion, "Note")
        
        If wMsg = vbYes Then
            TxtCode.Text = Replace(TxtCode.Text, "<?php", "")
            TxtCode.Text = Replace(TxtCode.Text, "?>", "")
        End If

    End If
    SttBar.Panels(1).Text = CDlog.FileName
    Me.Caption = "PHPcrypta - [" & CDlog.FileTitle & "]"
    TxtCode.SetFocus
End With



Err:
End Sub

Private Sub CmdEncrypt_Click()
On Error GoTo Err2:

Dim EnLvl As Integer
Dim TmpTxt As String
Dim EnText As String
Dim CdText As String
Dim DataFile As String
Dim CodeFile As String
Dim DataFileTitle As String
Dim FFnum As Integer
Dim vMSG

CdText = ""
EnText = ""
EnLvl = 0

OptLvl(0).Enabled = False
OptLvl(1).Enabled = False
OptLvl(2).Enabled = False

Mnu_lvl(0).Enabled = False
Mnu_lvl(1).Enabled = False
Mnu_lvl(2).Enabled = False

TxtPass(0).Enabled = False
TxtPass(1).Enabled = False
TxtPass(2).Enabled = False

TxtCode.Enabled = False
CmdBrowse.Enabled = False
CmdEncrypt.Enabled = False
Mnu_Start.Enabled = False
Mnu_New.Enabled = False
Mnu_Open.Enabled = False

If InStr(TxtCode.Text, "<?") > 0 Or InStr(TxtCode.Text, "?>") > 0 Then
    wMsg = MsgBox("We found '<?php' and '?>' tags, do you want to delete them ?", vbYesNo + vbQuestion, "Note")
    
    If wMsg = vbYes Then
        TxtCode.Text = Replace(TxtCode.Text, "<?php", "")
        TxtCode.Text = Replace(TxtCode.Text, "?>", "")
    End If

End If

If OptLvl(0).Value = True Then EnLvl = 0
If OptLvl(1).Value = True Then EnLvl = 1
If OptLvl(2).Value = True Then EnLvl = 2

ReSelectData:

With CDlog
    .FileName = "cryptaDATA.php"
    .DialogTitle = "Save Encypted Data File"
    .Filter = "(.php) PHP File|*.php"
    
    .CancelError = True
    .ShowSave
    DataFile = .FileName
    DataFileTitle = .FileTitle
    If Dir(DataFile) <> "" Then
        vMSG = MsgBox("Data File already exist, do you want to override it ?", vbQuestion + vbYesNoCancel, "File Exists")
        
        If vMSG = vbNo Then
            GoTo ReSelectData:
        ElseIf vMSG = vbCancel Then
            GoTo Err2:
        End If
    End If
End With

ReSelectCode:
With CDlog
    .FileName = "cryptaCODE.php"
    .DialogTitle = "Save Code File"
    .Filter = "(.php) PHP File|*.php"
    
    .CancelError = True
    .ShowSave
    CodeFile = .FileName
    
    If Dir(CodeFile) <> "" Then
        vMSG = MsgBox("Code File already exist, do you want to override it ?", vbQuestion + vbYesNoCancel, "File Exists")
        
        If vMSG = vbNo Then
            GoTo ReSelectCode:
        ElseIf vMSG = vbCancel Then
            GoTo Err2:
        End If
    End If
End With

If DataFile = CodeFile Then
    MsgBox "Data File can't be same as Code File, Please reselect code file.", vbCritical + vbOKOnly, "Error"
    GoTo ReSelectCode:
End If


On Error GoTo Err:
Select Case EnLvl
    Case 0
        
        For i = 1 To Len(TxtCode.Text)
            DoEvents
            PrgBar.Value = (i / Len(TxtCode.Text)) * PrgBar.Max
            SttBar.Panels(2).Text = Format(PrgBar.Value, "#") & "%"
            TmpTxt = Val(Asc(Mid(TxtCode.Text, i, 1))) + Asc("A") - Asc("y") + Asc("o") + Asc("o") + Asc("b")
            EnText = EnText & Format(TmpTxt, "000")
        Next i
        
        CdText = "foreach( str_split($cryptaDATA, 3) as $spData ) {" & vbCrLf _
        & "    $chrData = $spData - ord(""A"") + ord(""y"") - ord(""o"") - ord(""o"") - ord(""b"");" & vbCrLf _
        & "    $cryptaCode .= chr($chrData);" & vbCrLf _
        & "}" & vbCrLf
        
    Case 1
        
        For i = 1 To Len(TxtCode.Text)
            DoEvents
            PrgBar.Value = (i / Len(TxtCode.Text)) * PrgBar.Max
            SttBar.Panels(2).Text = Format(PrgBar.Value, "#") & "%"
            TmpTxt = ((Val(Asc(Mid(TxtCode.Text, i, 1))) + Val(TxtPass(0).Text)) * Val(TxtPass(1).Text)) _
            + Val(TxtPass(2).Text) - Val(TxtPass(1).Text) + Asc("A") - Asc("y") - Asc("o") + Asc("o") - Asc("b")
            EnText = EnText & Val(TmpTxt) & ";"
        Next i
        
        EnText = Left(EnText, Len(EnText) - 1)
        
        CdText = "foreach( explode("";"",$cryptaDATA) as $spData ) {" & vbCrLf _
        & "    $chrData = (($spData + ord(""b"") - ord(""o"") + ord(""o"") + ord(""y"") - ord(""A"") + intval(" & Val(TxtPass(1).Text) & ") - intval(" & Val(TxtPass(2).Text) & ")) / intval(" & Val(TxtPass(1).Text) & ")) - (intval(" & Val(TxtPass(0).Text) & "));" & vbCrLf _
        & "    $cryptaCode .= chr($chrData);" & vbCrLf _
        & "}" & vbCrLf
        
    Case 2
        
        For i = 1 To Len(TxtCode.Text)
            DoEvents
            PrgBar.Value = (i / Len(TxtCode.Text)) * PrgBar.Max
            SttBar.Panels(2).Text = Format(PrgBar.Value, "#") & "%"
            TmpTxt = ((Val(Asc(Mid(TxtCode.Text, i, 1))) + Val(TxtPass(0).Text)) * Val(TxtPass(1).Text)) _
            + Val(TxtPass(2).Text) + Val(TxtPass(0).Text) - Val(TxtPass(1).Text) + Asc("A") + Asc("y") - Asc("o") + Asc("o") - Asc("b")
            TmpTxt2 = ""
            For ii = 1 To Len(TmpTxt)
                TmpTxt2 = TmpTxt2 & Val(Mid(TmpTxt, ii, 1)) & Chr(Int((122 - 97) * Rnd + 97))
            Next ii
            
            TmpTxt = TmpTxt2
            
            EnText = EnText & TmpTxt & ";"
        Next i

        EnText = Left(EnText, Len(EnText) - 1)
        
        CdText = "for ($i = 97; $i <= 122; $i++) {" & vbCrLf _
        & "    $cryptaDATA = str_replace(chr($i), """", $cryptaDATA);" & vbCrLf _
        & "}" & vbCrLf & vbCrLf _
        & "foreach( explode("";"",$cryptaDATA) as $spData ) {" & vbCrLf _
        & "    $chrData = (($spData + ord(""b"") - ord(""o"") + ord(""o"") - ord(""y"") - ord(""A"") + intval(" & Val(TxtPass(1).Text) & ") - intval(" & Val(TxtPass(2).Text) & ") - intval(" & Val(TxtPass(0).Text) & ")) / intval(" & Val(TxtPass(1).Text) & ")) - (intval(" & Val(TxtPass(0).Text) & "));" & vbCrLf _
        & "    $cryptaCode .= chr($chrData);" & vbCrLf _
        & "}" & vbCrLf
        
End Select

FFnum = FreeFile
Open DataFile For Output As #FFnum
    
    Print #FFnum, "<?php"
    Print #FFnum, "/**"
    Print #FFnum, " * |================================|"
    Print #FFnum, " * |  File Encrypted By: PHPcrypta  |"
    Print #FFnum, " * |       www.AXEsystems.com       |"
    Print #FFnum, " * |     Developed By: AyoobAli     |"
    Print #FFnum, " * |         www.Atyus.com          |"
    Print #FFnum, " * |              2013              |"
    Print #FFnum, " * |================================|"
    Print #FFnum, " **/"
    Print #FFnum, ""
    Print #FFnum, "$cryptaDATA = """ & EnText & """;"
    Print #FFnum, ""
    Print #FFnum, "?>"
    
Close #FFnum

FFnum = FreeFile
Open CodeFile For Output As #FFnum
    
    Print #FFnum, "<?php"
    Print #FFnum, "/**"
    Print #FFnum, " * |================================|"
    Print #FFnum, " * |  File Encrypted By: PHPcrypta  |"
    Print #FFnum, " * |       www.AXEsystems.com       |"
    Print #FFnum, " * |     Developed By: AyoobAli     |"
    Print #FFnum, " * |         www.Atyus.com          |"
    Print #FFnum, " * |              2013              |"
    Print #FFnum, " * |================================|"
    Print #FFnum, " **/"
    Print #FFnum, ""
    Print #FFnum, "@include(""" & DataFileTitle & """);"
    Print #FFnum, ""
    Print #FFnum, "$cryptaCode = """";"
    Print #FFnum, ""
    Print #FFnum, CdText
    Print #FFnum, ""
    Print #FFnum, "@eval($cryptaCode);"
    Print #FFnum, ""
    Print #FFnum, "?>"
    
Close #FFnum

If Right(SttBar.Panels(1).Text, 1) = "*" Then SttBar.Panels(1).Text = Left(SttBar.Panels(1).Text, Len(SttBar.Panels(1).Text) - 1)

MsgBox "Encryption has been successfully saved.", vbOKOnly, "Done."

PrgBar.Value = 0
SttBar.Panels(2).Text = ""

OptLvl(0).Enabled = True
OptLvl(1).Enabled = True
OptLvl(2).Enabled = True

Mnu_lvl(0).Enabled = True
Mnu_lvl(1).Enabled = True
Mnu_lvl(2).Enabled = True

If OptLvl(0).Value = False Then TxtPass(0).Enabled = True
If OptLvl(0).Value = False Then TxtPass(1).Enabled = True
If OptLvl(0).Value = False Then TxtPass(2).Enabled = True

TxtCode.Enabled = True
CmdBrowse.Enabled = True
CmdEncrypt.Enabled = True
Mnu_Start.Enabled = True
Mnu_New.Enabled = True
Mnu_Open.Enabled = True


Exit Sub
Err:
MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"

Err2:
OptLvl(0).Enabled = True
OptLvl(1).Enabled = True
OptLvl(2).Enabled = True

Mnu_lvl(0).Enabled = True
Mnu_lvl(1).Enabled = True
Mnu_lvl(2).Enabled = True

If OptLvl(0).Value = False Then TxtPass(0).Enabled = True
If OptLvl(0).Value = False Then TxtPass(1).Enabled = True
If OptLvl(0).Value = False Then TxtPass(2).Enabled = True

TxtCode.Enabled = True
CmdBrowse.Enabled = True
CmdEncrypt.Enabled = True
Mnu_Start.Enabled = True
Mnu_New.Enabled = True
Mnu_Open.Enabled = True
End Sub

Private Sub Form_Resize()
On Error Resume Next

    If Me.Width < 7050 Then Me.Width = 7050
    If Me.Height < 2900 Then Me.Height = 2900
    
    ShpBG(1).Width = Me.Width - 5160
    ShpBG(3).Width = Me.Width - 5640
    ShpBG(4).Width = Me.Width
    ShpBG(4).Height = Me.Height - 2070
    
    TxtCode.Width = Me.Width - 480
    TxtCode.Height = Me.Height - 2550
    
    PrgBar.Width = Me.Width - 6960
    TxtFile.Width = Me.Width - 6840
    
    
    CmdBrowse.Left = Me.Width - 1455
    CmdEncrypt.Left = Me.Width - 1695

End Sub

Private Sub Mnu_About_Click()
On Error Resume Next
    FrmAbout.Show 1, Me
End Sub

Private Sub Mnu_Exit_Click()
On Error Resume Next
    Unload FrmAbout
    Unload Me
    End
End Sub

Private Sub Mnu_Instructions_Click()
On Error Resume Next
    ShellExecute hwnd, "open", "http://www.axesystems.com/software/phpcrypta/", vbNullString, vbNullString, Empty
End Sub

Private Sub Mnu_lvl_Click(Index As Integer)
On Error Resume Next

    Mnu_lvl(0).Checked = False
    Mnu_lvl(1).Checked = False
    Mnu_lvl(2).Checked = False
    Mnu_lvl(Index).Checked = True
    
    OptLvl(Index).Value = True
    OptLvl_Click (Index)
    
End Sub

Private Sub Mnu_New_Click()
On Error GoTo Err:

Dim ConfirmMSG

If Right(SttBar.Panels(1).Text, 1) = "*" Then
    ConfirmMSG = MsgBox("Are you sure you want to start a new document ?" & vbCrLf & "All Changes will be lost.", vbQuestion + vbYesNo, "Open New File")
    
    If ConfirmMSG = vbNo Then GoTo Err:
End If

TxtCode.Text = ""
Mnu_lvl_Click (0)
TxtFile.Text = ""
SttBar.Panels(1).Text = "Select File..."
Me.Caption = "PHPcrypta"

Err:

End Sub

Private Sub Mnu_Open_Click()
On Error Resume Next
    CmdBrowse_Click
End Sub



Private Sub Mnu_Start_Click()
On Error Resume Next
    CmdEncrypt_Click
End Sub

Private Sub OptLvl_Click(Index As Integer)
On Error GoTo Err:

Mnu_lvl(0).Checked = False
Mnu_lvl(1).Checked = False
Mnu_lvl(2).Checked = False
Mnu_lvl(Index).Checked = True
Select Case Index
    Case 0
        LblPass.Enabled = False
        TxtPass(0).Enabled = False
        TxtPass(1).Enabled = False
        TxtPass(2).Enabled = False

    Case 1
        LblPass.Enabled = True
        TxtPass(0).Enabled = True
        TxtPass(1).Enabled = True
        TxtPass(2).Enabled = True

    Case 2
        LblPass.Enabled = True
        TxtPass(0).Enabled = True
        TxtPass(1).Enabled = True
        TxtPass(2).Enabled = True

End Select

Err:
End Sub

Private Sub TxtCode_Change()
On Error Resume Next
    If Me.Caption = "PHPcrypta" And Trim(TxtCode.Text) = "" Then
        SttBar.Panels(1).Text = "Select File..."
        CmdEncrypt.Enabled = False
    Else
        If Right(SttBar.Panels(1).Text, 1) <> "*" Then SttBar.Panels(1).Text = SttBar.Panels(1).Text & "*"
        CmdEncrypt.Enabled = True
    End If
End Sub


Private Sub TxtPass_Change(Index As Integer)
On Error Resume Next
    
    If Val(TxtPass(Index).Text) < 1 Then
        TxtPass(Index).Text = 1
        TxtPass(Index).SelStart = 0
        TxtPass(Index).SelLength = Len(TxtPass(Index).Text)
    End If
    
End Sub

Private Sub TxtPass_GotFocus(Index As Integer)
On Error Resume Next
    TxtPass(Index).SelStart = 0
    TxtPass(Index).SelLength = Len(TxtPass(Index).Text)
    
End Sub

Private Sub TxtPass_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtPass_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If (KeyCode = 13 Or Len(TxtPass(Index).Text) > 2) And Index < 2 Then
        TxtPass(Index + Val(1)).SetFocus
    End If
End Sub
