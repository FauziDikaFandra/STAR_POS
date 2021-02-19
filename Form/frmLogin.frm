VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1545
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   7995
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1545
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   7995
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4275
      Left            =   75
      TabIndex        =   19
      Top             =   1950
      Width           =   4215
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2000
         Left            =   0
         ScaleHeight     =   1965
         ScaleWidth      =   5970
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   6000
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   13
         Left            =   3075
         TabIndex        =   13
         Top             =   3150
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
         TabIndex        =   12
         Top             =   1725
         Width           =   975
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
         TabIndex        =   11
         Top             =   225
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
         TabIndex        =   14
         Top             =   3150
         Width           =   1950
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
         TabIndex        =   1
         Top             =   3150
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   2175
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
         TabIndex        =   4
         Top             =   2175
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   1200
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
         TabIndex        =   7
         Top             =   1200
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
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   225
         Width           =   975
      End
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
         TabIndex        =   10
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   75
      TabIndex        =   16
      Top             =   525
      Width           =   4215
      Begin TDBText6Ctl.TDBText txtuser 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   900
         Width           =   3840
         _Version        =   65536
         _ExtentX        =   6773
         _ExtentY        =   688
         Caption         =   "frmLogin.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLogin.frx":0934
         Key             =   "frmLogin.frx":0952
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
      Begin TDBText6Ctl.TDBText txtregid 
         Height          =   390
         Left            =   150
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   225
         Width           =   2190
         _Version        =   65536
         _ExtentX        =   3863
         _ExtentY        =   688
         Caption         =   "frmLogin.frx":0996
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLogin.frx":0A08
         Key             =   "frmLogin.frx":0A26
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   "888"
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
         TabIndex        =   15
         Top             =   900
         Visible         =   0   'False
         Width           =   3840
         _Version        =   65536
         _ExtentX        =   6773
         _ExtentY        =   688
         Caption         =   "frmLogin.frx":0A6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLogin.frx":0AD6
         Key             =   "frmLogin.frx":0AF4
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
      Begin VB.Label lblonline 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2475
         TabIndex        =   20
         Top             =   225
         Width           =   1515
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4200
         Y1              =   750
         Y2              =   750
      End
   End
   Begin VB.Label lbltoko 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "STAR STAR STAR STAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ulang As Byte
Dim posisi As String







Private Sub Form_Activate()
'    txtuser = "1313"
'    txtpassword = "2011"
'    Call cmdLogin_Click
End Sub

Private Sub Form_Load()
    lbltoko = Tulis(10)
    txtregid = VReg_ID
    lblonline = VPing

End Sub

Private Sub btnNum_Click(Index As Integer)
    Select Case posisi
    Case "txtuser"
        If Index < 10 Then txtuser.Text = txtuser.Text + btnNum(Index).Caption
        
        Select Case Index
        Case 10
            txtuser.SetFocus
            SendKeys "{end}+{backspace}"
        Case 11
            txtuser.SetFocus
            SendKeys "{enter}"
        Case 12
            txtuser.Text = ""
        Case 13
            End
        End Select
          
    Case "txtpassword"
        If Index < 10 Then txtpassword.Text = txtpassword.Text + btnNum(Index).Caption
        
        Select Case Index
        Case 10
            txtpassword.SetFocus
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
            End
        End Select
        
    End Select
End Sub

Private Sub cmdLogin_Click()
Dim RsUser As New ADODB.Recordset
    StrSQL = "select * from users where User_ID='" & txtuser.Text & "' and branch_id='" & VBranch_ID & "'"
   
    If Linked Then
        RsUser.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsUser.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
        
    If Not RsUser.EOF Then
        If Trim(RsUser!Password) = txtpassword.Text Then
            VKasir_ID = txtuser.Text
            VKasir_Nm = RsUser!user_Name
            
            Unload Me
            DoEvents
            
            If Linked Then Call Cek_SOD("server")
            Call Cek_SOD("local")
            Call Cek_Card
            Call Cek_CashOpen
            Call Cek_AdaPromo
            Call Isi_key
            Call CetakBegin
            frmMain.Show
            
            Call SQLQuery("update cash_register set active_status=1 where Branch_ID = '" & _
                        VBranch_ID & "' AND cash_register_id='" & VReg_ID & "'")
            
            Call SaveLog("Login Success" & " " & VKasir_ID & " / " & VKasir_Nm)
        Else
            MsgBox "Password yang anda masukkan salah", vbCritical + vbOKOnly, "Oops.."
            ulang = ulang + 1
            If ulang > 2 Then End
            txtuser.Visible = False
            txtpassword.Visible = True
            txtpassword.SetFocus
            SendKeys "{home}+{end}"
      End If
    Else
        MsgBox "Username tidak terdaftar", vbCritical + vbOKOnly, "Oops.."
        txtuser.Visible = True
        txtpassword.Visible = False
        txtuser.SetFocus
        SendKeys "{home}+{end}"
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
        txtuser.Visible = True
        txtpassword.Visible = False
        txtuser.SetFocus
        txtpassword.Text = ""
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
        End
    End Select
End Sub

Private Sub Cek_SOD(StrSvr As String)
Dim RsSOD As New ADODB.Recordset

    StrSQL = "select flag_sod from branches where branch_id ='" & VBranch_ID & "'"
    
    Select Case StrSvr
    Case "server"
        RsSOD.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
        If RsSOD!Flag_SOD = 0 Then
            MsgBox "Server belum SOD..", vbCritical + vbOKOnly, "Oops.."
            End
        End If

    Case "local"
        RsSOD.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
        If VPing = "ONLINE" And RsSOD!Flag_SOD = 0 Then
            frmSOD.Caption = "SOD"
            frmSOD.Show 1
        End If
    End Select
    RsSOD.Close: Set RsSOD = Nothing
End Sub

Private Sub Cek_Card()
Dim RsCcard As New ADODB.Recordset

    RsCcard.Open "select count(*) as berapa from customer_master_member", ConnLocal, adOpenForwardOnly, adLockReadOnly
    If RsCcard!Berapa = 0 Then
       ConnLocal.Execute ("update branches set flag_sod=0 where Branch_ID = '" & VBranch_ID & "'")
       MsgBox "SOD belum selesai, jalankan iPOS sekali lagi...", vbCritical + vbOKOnly, "Oops.."
       End
    End If
    RsCcard.Close: Set RsCcard = Nothing
End Sub

Private Sub Cek_CashOpen()
Dim RsCash As New ADODB.Recordset
        
    StrSQL = "select modal from cash where branch_id = '" & VBranch_ID & "' and Cash_Register_ID='" & _
                VReg_ID & " ' and shift='" & VShift & "' and datetime='" & Format(GetSrvDate, "YYYY-MM-DD") & "'"
   
    If Linked Then
        RsCash.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsCash.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
 
    VCopen = Not RsCash.EOF
    RsCash.Close: Set RsCash = Nothing
End Sub

Private Sub Cek_AdaPromo()
Dim RsPro As New ADODB.Recordset
    
    RsPro.Open "select promo_id from promo_hdr where getdate() Between Start_Date And End_Date And aktif=1", _
               ConnLocal, adOpenForwardOnly, adLockReadOnly
               
    VAda_Promo = Not RsPro.EOF
    RsPro.Close
    
    RsPro.Open "Select * from informasi", ConnLocal, adOpenForwardOnly, adLockReadOnly
    If Not RsPro.EOF Then
        Tulis(16) = RsPro!Pesan1 & vbNewLine & RsPro!Pesan2 & vbNewLine & _
                    RsPro!Pesan3 & vbNewLine & RsPro!Pesan4 & vbNewLine & _
                    RsPro!Pesan5 & vbNewLine & RsPro!Pesan6 & vbNewLine & _
                    RsPro!pesan7 & vbNewLine & RsPro!pesan8
    End If
    RsPro.Close:   Set RsPro = Nothing
End Sub

Private Sub Isi_key()
Dim RsIsi As New ADODB.Recordset

    RsIsi.Open "select form, menu, keycode from key_map", ConnLocal, adOpenForwardOnly, adLockReadOnly
    While Not RsIsi.EOF
        If RsIsi!KeyCode < 122 And RsIsi!KeyCode > 26 Then KeyStroke(RsIsi!KeyCode) = RsIsi!Menu
        RsIsi.MoveNext
    Wend
    RsIsi.Close: Set RsIsi = Nothing
End Sub
