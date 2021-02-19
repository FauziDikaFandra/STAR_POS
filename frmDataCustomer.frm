VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmDataCustomer 
   Caption         =   "frmDataCustomer"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3000
      Picture         =   "frmDataCustomer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1650
      Picture         =   "frmDataCustomer.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6210
      Begin TDBText6Ctl.TDBText txtname 
         Height          =   390
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   4950
         _Version        =   65536
         _ExtentX        =   8731
         _ExtentY        =   688
         Caption         =   "frmDataCustomer.frx":1404
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDataCustomer.frx":1468
         Key             =   "frmDataCustomer.frx":1486
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   50
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
      Begin TDBText6Ctl.TDBText txtemail 
         Height          =   390
         Left            =   150
         TabIndex        =   4
         Top             =   1200
         Width           =   4950
         _Version        =   65536
         _ExtentX        =   8731
         _ExtentY        =   688
         Caption         =   "frmDataCustomer.frx":14CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDataCustomer.frx":1530
         Key             =   "frmDataCustomer.frx":154E
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   50
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
      Begin TDBText6Ctl.TDBText txtno_telp 
         Height          =   390
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   4290
         _Version        =   65536
         _ExtentX        =   7567
         _ExtentY        =   688
         Caption         =   "frmDataCustomer.frx":15A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDataCustomer.frx":160A
         Key             =   "frmDataCustomer.frx":1628
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
         MaxLength       =   15
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
      Begin VB.Label txtcapt 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label txtcardx 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label vtotalx 
         Caption         =   "Label3"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label vdiscx 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Vpromo_id 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   3
         Top             =   3480
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5550
      Top             =   750
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "DATA CUSTOMER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   450
      Left            =   0
      TabIndex        =   7
      Top             =   15
      Width           =   5400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HARAP DITANYAKAN DATA CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   225
      TabIndex        =   5
      Top             =   4545
      Width           =   1665
   End
End
Attribute VB_Name = "frmDataCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cmdcancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If Trim(Star_Id) = "100000" Then
        txtname.SetFocus
    Else
        txtname = Star_Nm
        txtno_telp = Star_Phone
        txtemail = Star_Email
        txtemail.SetFocus
    End If
End Sub



Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtemail = "" Then
            Exit Sub
        End If
    CekDataEmail (txtemail)
    CmdOk.SetFocus
    End If
End Sub

Private Function CekDataEmail(Email As String) As String
Dim RsCari As New ADODB.Recordset
        StrSQL = "select * from EReceipt_Email_Contact where Email = '" & Email & "'"
        If Linked Then
            RsCari.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
        Else
            RsCari.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
        End If
        CekDataEmail = ""
        If Not RsCari.EOF Then
            CekDataEmail = RsCari!Email
            txtname = RsCari!Nama
            txtno_telp = RsCari!Hp
        End If
End Function

Private Sub Cmdok_Click()

        If txtname = "" Then
            MsgBox ("Nama di isi terlebih dahulu !!?")
            txtname.SetFocus
            Exit Sub
        End If
        If Len(txtno_telp) < 10 Then
            MsgBox ("No Telp Tidak Valid Coba ulangi !!?")
            txtno_telp.SetFocus
            Exit Sub
        End If
            If isValidEmail(txtemail) = False Then
            MsgBox ("Email Tidak Valid Coba ulangi !!?")
            txtemail.SetFocus
            Exit Sub
        End If
        InputNama = txtname
        InputNoTlp = txtno_telp
        InputEmail = txtemail
        With frmPayment
            Call CDisplay("TOTAL :", "Rp. " & Format(vtotalx - vdiscx, "#,##0"))
            .vpay = vtotalx - vdiscx
            .txtcard_no = txtcardx
            .vstatus = txtcapt
            .Show 1
        End With
        Unload Me
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If txtname = "" Then
        Exit Sub
    End If
    txtno_telp.SetFocus
End If
End Sub

Private Sub txtno_telp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If txtno_telp = "" Then
        Exit Sub
    End If
    txtemail.SetFocus
Exit Sub
End If
End Sub


