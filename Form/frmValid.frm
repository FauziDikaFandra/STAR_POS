VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmValid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VERIFICATION"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmValid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   16
      Top             =   75
      Width           =   4215
      Begin TDBText6Ctl.TDBText txtuser 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   150
         Width           =   3840
         _Version        =   65536
         _ExtentX        =   6773
         _ExtentY        =   688
         Caption         =   "frmValid.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValid.frx":0934
         Key             =   "frmValid.frx":0952
         BackColor       =   16777215
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin TDBText6Ctl.TDBText txtpassword 
         Height          =   390
         Left            =   150
         TabIndex        =   17
         Top             =   150
         Visible         =   0   'False
         Width           =   3840
         _Version        =   65536
         _ExtentX        =   6773
         _ExtentY        =   688
         Caption         =   "frmValid.frx":0996
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValid.frx":0A02
         Key             =   "frmValid.frx":0A20
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   "*"
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   6
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   3
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   75
      TabIndex        =   15
      Top             =   825
      Width           =   4215
      Begin VB.CommandButton btnNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   2100
         TabIndex        =   11
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   1125
         TabIndex        =   10
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   150
         TabIndex        =   9
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   2100
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   1125
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   2100
         TabIndex        =   5
         Top             =   2175
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1125
         TabIndex        =   4
         Top             =   2175
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   2175
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   3150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "ENTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   1125
         TabIndex        =   1
         Top             =   3150
         Width           =   1950
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Index           =   10
         Left            =   3075
         TabIndex        =   12
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Index           =   12
         Left            =   3075
         TabIndex        =   13
         Top             =   1725
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   13
         Left            =   3075
         TabIndex        =   14
         Top             =   3150
         Width           =   975
      End
   End
   Begin VB.Label VLevelApp 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4425
      TabIndex        =   18
      Top             =   900
      Width           =   1515
   End
End
Attribute VB_Name = "frmValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ulang As Byte
Dim posisi As String

Private Sub btnNum_Click(Index As Integer)
    Select Case posisi
    Case "txtuser"
        txtuser.SetFocus
        If Index < 10 Then txtuser.Text = txtuser.Text + btnNum(Index).Caption
        
        Select Case Index
        Case 10
            SendKeys "{end}+{backspace}"
        Case 11
            SendKeys "{enter}"
        Case 12
            txtuser.Text = ""
        Case 13
            Unload Me
        End Select
          
    Case "txtpassword"
        txtpassword.SetFocus
        If Index < 10 Then txtpassword.Text = txtpassword.Text + btnNum(Index).Caption
        
        Select Case Index
        Case 10
            SendKeys "{end}+{backspace}"
        Case 11
            cmdLogin_Click
        Case 12
            If txtpassword.Text = "" Then
                txtuser.Visible = True
                txtpassword.Visible = False
                txtuser.SetFocus
            End If
            txtpassword.Text = ""
        Case 13
            Unload Me
        End Select
    End Select
End Sub

Private Sub cmdLogin_Click()
Dim RsUser As New ADODB.Recordset
    StrSQL = "select * from users where User_ID='" & txtuser.Text & _
             "' and branch_id='" & VBranch_ID & "' and security_level <='" & VLevelApp & "'"
   
    If Linked Then
        RsUser.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsUser.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
        
    If Not RsUser.EOF Then
        If Trim(RsUser!Password) = txtpassword.Text Then
           VSuper_Nm = RsUser!user_id & " / " & RsUser!user_Name
           VOK = True
           Call SaveLog("Verifikasi Success -" & VLevelApp & " " & RsUser!user_id & " / " & RsUser!user_Name)
           Unload Me
        Else
            MsgBox "Password yang anda masukkan salah", vbCritical + vbOKOnly, "Oops.."
            ulang = ulang + 1
            If ulang > 2 Then
                Unload Me
                ulang = 0
                Exit Sub
            End If
            txtpassword.SetFocus
            SendKeys "{home}+{end}"
        End If
   Else
        MsgBox "User tidak ada otorisasi", vbCritical + vbOKOnly, "Oops.."
        Unload Me
   End If
   RsUser.Close: Set RsUser = Nothing
End Sub

Private Sub txtpassword_GotFocus()
    posisi = "txtpassword"
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        cmdLogin_Click
    Case 27
        Unload Me
    End Select
End Sub

Private Sub txtuser_GotFocus()
    posisi = "txtuser"
End Sub

Private Sub txtuser_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        txtuser.Visible = False
        txtpassword.Visible = True
        txtpassword = ""
        txtpassword.SetFocus
    Case 27
        Unload Me
    End Select
End Sub
