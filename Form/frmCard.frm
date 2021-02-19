VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmCard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MYSTAR CARD"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   ControlBox      =   0   'False
   Icon            =   "frmCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNav 
      Caption         =   "&PHONE"
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
      Index           =   3
      Left            =   4650
      Picture         =   "frmCard.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4500
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
      Left            =   2085
      Picture         =   "frmCard.frx":28D3
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   1140
   End
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
      Left            =   3345
      Picture         =   "frmCard.frx":32D5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5550
      Top             =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3870
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   6090
      Begin TDBText6Ctl.TDBText txtPeriod 
         Height          =   390
         Left            =   3300
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   688
         Caption         =   "frmCard.frx":3CD7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":3D33
         Key             =   "frmCard.frx":3D51
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
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
         MaxLength       =   0
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
      Begin TDBText6Ctl.TDBText txtexprPoint 
         Height          =   375
         Left            =   150
         TabIndex        =   13
         Top             =   3360
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   661
         Caption         =   "frmCard.frx":3D95
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":3E0B
         Key             =   "frmCard.frx":3E29
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
         MaxLength       =   5
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
      Begin TDBText6Ctl.TDBText txtcard_no 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   225
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   688
         Caption         =   "frmCard.frx":3E6D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":3EE5
         Key             =   "frmCard.frx":3F03
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
         MaxLength       =   12
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
      Begin TDBText6Ctl.TDBText txtcust_name 
         Height          =   390
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   750
         Width           =   5790
         _Version        =   65536
         _ExtentX        =   10213
         _ExtentY        =   688
         Caption         =   "frmCard.frx":3F47
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":3FBD
         Key             =   "frmCard.frx":3FDB
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
         MaxLength       =   0
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
      Begin TDBText6Ctl.TDBText txtcard_opt 
         Height          =   390
         Left            =   150
         TabIndex        =   3
         Top             =   2850
         Width           =   3915
         _Version        =   65536
         _ExtentX        =   6906
         _ExtentY        =   688
         Caption         =   "frmCard.frx":402D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":409F
         Key             =   "frmCard.frx":40BD
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
         MaxLength       =   16
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
      Begin TDBText6Ctl.TDBText txtpoint 
         Height          =   390
         Left            =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   688
         Caption         =   "frmCard.frx":4101
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":4175
         Key             =   "frmCard.frx":4193
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
         MaxLength       =   5
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
      Begin TDBText6Ctl.TDBText txtcust_id 
         Height          =   390
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1275
         Width           =   5790
         _Version        =   65536
         _ExtentX        =   10213
         _ExtentY        =   688
         Caption         =   "frmCard.frx":41E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":425D
         Key             =   "frmCard.frx":427B
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
         MaxLength       =   0
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
      Begin TDBText6Ctl.TDBText txtfreq 
         Height          =   390
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   688
         Caption         =   "frmCard.frx":42CD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":4341
         Key             =   "frmCard.frx":435F
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
         MaxLength       =   5
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
         Left            =   2970
         TabIndex        =   15
         Top             =   1800
         Width           =   2970
         _Version        =   65536
         _ExtentX        =   5239
         _ExtentY        =   688
         Caption         =   "frmCard.frx":43B1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":441B
         Key             =   "frmCard.frx":4439
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
      Begin TDBText6Ctl.TDBText txt_email 
         Height          =   390
         Left            =   150
         TabIndex        =   16
         Top             =   2295
         Width           =   5770
         _Version        =   65536
         _ExtentX        =   10178
         _ExtentY        =   688
         Caption         =   "frmCard.frx":447D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCard.frx":44E3
         Key             =   "frmCard.frx":4501
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
         MaxLength       =   40
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
         Left            =   4425
         TabIndex        =   7
         Top             =   2850
         Width           =   1515
      End
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
      TabIndex        =   11
      Top             =   4545
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   765
      Left            =   150
      Top             =   4470
      Width           =   1815
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "SCAN BARCODE KARTU"
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
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6240
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub CmdNav_Click(Index As Integer)
    frmNum.Caption = "NUMBER - TELP"
    frmNum.Show 1
    txtno_telp.SetFocus
End Sub

Private Sub Cmdok_Click()
    If txtcard_no = "" Then
        Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
        Star_Id = "100000"
        Exit Sub
    End If
    
    If txtcard_opt <> "" Then Call txtcard_opt_KeyDown(13, 0)
    
    VBonus_Point = 1
    frmSales.Caption = Me.Caption
    frmSales.txtcard_no = txtcard_no
    frmSales.txtcust_name = txtcust_name
    frmSales.txtpoint = txtpoint
    frmSales.txtcust_id = txtcust_id
    Star_No = txtcard_no
    Star_Nm = txtcust_name
    Star_Phone = txtno_telp
    
    If txtcard_no <> "CM000-00000" Then
       Call Cek_Bonus_Point
       Call CDisplay(UCase(txtcard_no), Left(txtcust_name, 20))
    End If
    
    Unload Me
    Unload frmMain
    frmSales.Show
End Sub

Private Sub cmdok_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Call isi_data("", "", "", "", "", "", "", "", "")
    txtcard_no.Text = ""
    txtno_telp.Text = ""
    VIsSSC = False
    VIsKKG = False
    StrukEmail = False
    Dim sttStrukEmail As Integer
    sttStrukEmail = 0
    If MsgBox("Apakah Customer ingin struk dikirim via email?", vbYesNo + vbOKOnly, "Informasi") = vbYes Then
        Call CetakDataEmail
        StrukEmail = True
        sttStrukEmail = 1
    End If
    
    If VNomor <> "" Then
        ConnLocal.Execute "update sales_transactions set Payment_Program_ID = '" & sttStrukEmail & "' where Transaction_number = '" & VNomor & "'"
    End If
End Sub

Private Sub Timer1_Timer()
    lblmsg.ForeColor = IIf(lblmsg.ForeColor = vbBlue, vbYellow, vbBlue)
    If Len(txtcard_no.Text) <> 11 Then txtcard_no = ""
End Sub

Private Sub txtcard_no_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Shape1.BackColor = vbWhite
        txtcard_no = UCase(txtcard_no)
        
        If Left(txtcard_no.Text, 2) <> "CM" Then
            Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
            Star_Id = "100000"
            Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
            Star_Id = "100000"
            CmdOk.SetFocus
            Exit Sub
        End If
        
        If Len(txtcard_no.Text) > 11 And Mid(txtcard_no.Text, 12, 1) = "A" Then
            txtcard_no.Text = Left(txtcard_no.Text, 11)
            ScanApps = True
        Else
            ScanApps = False
        End If
        
        Call SQLQuery("update Card set Card_Status = 'A' where Card_Nr = '" & txtcard_no & "' and Card_Status = 'D'")
        Call MySTAR(txtcard_no, 0)
        If Star_Id = "100000" Then
            Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
            Star_Id = "100000"
        Else
            MSCTlp = False
            Call isi_data(txtcard_no, Star_Nm, CStr(Star_Pt), Star_Id, Star_Freq, Star_Email, Star_Phone, Exp_Point, Expired_Date)
            If Linked And Star_updsts = 0 Then
                Call CDisplay(Star_Phone, Left(Star_Email, 20))
                Shape1.BackColor = vbBlue
                Star_No = txtcard_no
                If MsgBox("Apakah Customer ingin mengupdate data?" & vbNewLine & _
                    "Nama :  " & Star_Nm & vbNewLine & _
                    "Phone :  " & Star_Phone & vbNewLine & _
                    "Email   :  " & Star_Email, vbYesNo + vbOKOnly, "Update Data Member") = vbYes Then
                    Call CetakData
                    Call SQLQuery("update card set update_status=2 where card_nr = '" & Star_No & "'")
                Else
                    Call SQLQuery("update card set update_status=1 where card_nr = '" & Star_No & "'")
                End If
            End If
            If Mid(txtcard_no, 1, 5) = "CM999" And Linked Then
            If Not IsNumeric(Star_Ext1) Then
                'MsgBox "Format potongan karyawan salah " & vbNewLine & _
                "Harap hubungi IT. ", vbCritical + vbOKOnly, "Oops.."
                'Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "")
                'Star_Id = "100000"
                'CmdOk.SetFocus
                'Exit Sub
            Else
            If Star_Ext1 > 10 Then
            MsgBox ("Sisa potongan karyawan anda senilai Rp." + FormatNumber(Star_Ext1, 0, True, True, True) & " !!!")
            End If
            End If
            End If
        End If
        txtcard_opt.SetFocus
    Case 27
        Unload Me
    End Select
End Sub

Private Sub isi_data(No_Kartu As String, Nama As String, Point As String, id As String, freq As String, Email As String, telepon As String, PointEx As String, PeriodEx As String)
    If telepon <> "" Then
      If Left(Star_Phone, 1) = "0" Then
        Star_Phone = Star_Phone
        telepon = Star_Phone
      Else
        Star_Phone = "0" & Mid(Star_Phone, 3, Len(Star_Phone))
        telepon = Star_Phone
      End If
    End If

                
    txtcard_no = No_Kartu
    txtcust_name = Nama
    txtpoint = Point
    txtcust_id = id
    txtfreq = freq
'    txtomz = Format(omz, "#,##0")
    txtno_telp = telepon
    txtexprPoint = PointEx
    txt_email = Email
    If Point <= PointEx Then
    txtexprPoint = 0
    End If
    txtPeriod = Format(PeriodEx, "DD-MMM-YYYY")
End Sub

Private Sub txtcard_opt_GotFocus()
    If txtcard_no = "CM000-00000" Then CmdOk.SetFocus
End Sub

Private Sub txtcard_opt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    Select Case KeyCode
    Case 13
        If Left(txtcard_opt, 2) = "CM" Or Len(txtcard_opt) < 3 Or Len(txtcard_opt) = 13 Then txtcard_opt = ""
        If txtcard_no = "CM000-00000" Or txtcard_opt = "" Then
            CmdOk.SetFocus
            Exit Sub
        End If
        
        Dim RsCari As New ADODB.Recordset
        StrSQL = "select * from card_promotion cp inner join card_promotion_name cn on cp.card_promo_id = cn.card_promo_id " & _
                 "where card_nr='" & txtcard_no & "' and card_nr_promo = '" & txtcard_opt & "'"

        If Linked Then
            RsCari.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
        Else
            RsCari.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
        End If

        If Not RsCari.EOF Then
            Select Case RsCari!card_promo_id
                Case "CPN001"
                    If GetSrvDate >= RsCari!start_promo_date And GetSrvDate <= RsCari!end_promo_date Then
                        'Promo SSC hanya untuk weekday saja
                        If (RsCari!card_promo_id = "CPN001" And Weekday(GetSrvDate) = 6) _
                        Or (RsCari!card_promo_id = "CPN001" And Weekday(GetSrvDate) = 7) _
                        Or (RsCari!card_promo_id = "CPN001" And Weekday(GetSrvDate) = 1) Then
                            MsgBox "Promo hanya berlaku weekday (Senin-Kamis)", vbInformation + vbOKOnly, "Oops.."
                            VBonus_Point = 1
                        Else
                            VBonus_Point = RsCari!point_bonus
                        End If
                        VIsSSC = True
                    Else
                        VBonus_Point = 1
                    End If
                    MsgBox "Bonus Point = " & VBonus_Point, vbInformation + vbOKOnly, "Oops.."
                    Call Cmdok_Click
                    Exit Sub
                Case "CPN002"
                    If GetSrvDate >= RsCari!start_promo_date And GetSrvDate <= RsCari!end_promo_date Then
                        VIsKKG = True
                    End If
            End Select
        Else
            If MsgBox("Daftarkan kartu promo baru?", vbQuestion + vbOKCancel, "Oops..") = vbOK Then
                frmCardPromo.Show 1
                DoEvents
                If Vpromo_id <> "" Then
                
                Dim RsDobel As New ADODB.Recordset
                    StrSQL = "select * from card_promotion where card_nr_promo ='" & txtcard_opt & "' and card_promo_id = '" & Vpromo_id & "'"
                    
                    If Linked Then
                        RsDobel.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
                    Else
                        RsDobel.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
                    End If
                
                    If Not RsDobel.EOF Then
                        MsgBox "Kartu joint promo sudah pernah ada, harap hubungi information counter", vbCritical + vbOKOnly, "Oops.."
                        Exit Sub
                    End If
                    RsDobel.Close: Set RsDobel = Nothing
                    
                    StrSQL = ("insert into Card_Promotion (Card_Nr, Card_Nr_Promo, Card_Promo_Id, Card_Expired_date, " & _
                            "Card_Activate_Date, Card_Status, User_Id_Activate) values ('" & txtcard_no & "','" & _
                            txtcard_opt & "','" & Vpromo_id & "',getdate()+1825, getdate(), 'A', '" & VKasir_ID & "')")
                    
                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
            End If
            txtcard_opt.SetFocus
            DoEvents
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        RsCari.Close: Set RsCari = Nothing
        CmdOk.SetFocus
    Case 27
        Unload Me
    End Select
    Exit Sub
    
ErrH:
    If Err.Number = "-2147217873" Then
        MsgBox "Kartu promo sudah pernah ada, harap hubungi information counter", vbCritical + vbOKOnly, "Oops.."
    Else
        MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    End If
    Call SaveLog(Me.name & " " & "Simpan Kartu Promo " & Err.Description & " " & Err.Number)
End Sub

Private Sub Cek_Bonus_Point()
Dim RsPoint As New ADODB.Recordset
Dim Hari As Byte

    Hari = IIf(Weekday(GetSrvDate) = 1, 7, Weekday(GetSrvDate) - 1)
    
    RsPoint.Open "select isnull(point,0) as point, substring(activeday," & Hari & ",1) as act_day " & _
                 "from cust_param_bonus where jenis_kartu='CM' and status_active='1' and GETDATE() between Start and Finish and substring(activeday," & Hari & ",1)='1'", _
                 ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If Not RsPoint.EOF Then
        VBonus_Point = IIf(VBonus_Point < RsPoint!Point, RsPoint!Point, VBonus_Point)
    End If
    RsPoint.Close: Set RsPoint = Nothing
End Sub

Private Sub txtno_telp_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        Shape1.BackColor = vbWhite
        txtno_telp = UCase(txtno_telp)
        If txtno_telp.Text = "" Then
            Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
            Star_Id = "100000"
            CmdOk.SetFocus
            Exit Sub
        End If
        
        
        
        If Len(Trim(txtno_telp.Text)) < 11 Then
            MsgBox ("No handphone yang di Input tidak sesuai !!!")
            txtno_telp.SetFocus
            Exit Sub
        End If
        
        If Left(txtno_telp.Text, 1) = "0" Then
            txtno_telp.Text = "62" & Mid(txtno_telp.Text, 2, Len(txtno_telp.Text))
        End If
               
        ScanApps = False
        
        Call MySTAR(txtcard_no, txtno_telp)
        If Star_Id = "100000" Then
            Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "", "", "")
            Star_Id = "100000"
        Else
            MSCTlp = True
            Call isi_data(txtcard_no, Star_Nm, CStr(Star_Pt), Star_Id, Star_Freq, Star_Email, Star_Phone, Exp_Point, Expired_Date)
            If Linked And Star_updsts = 0 Then
                Call CDisplay(Star_Phone, Left(Star_Email, 20))
                Shape1.BackColor = vbBlue
                Star_No = txtcard_no
                
                If MsgBox("Apakah Customer ingin mengupdate data?" & vbNewLine & _
                    "Phone :  " & Star_Phone & vbNewLine & _
                    "Email   :  " & Star_Email, vbYesNo + vbOKOnly, "Update Data Member") = vbYes Then
                    Call CetakData
                    Call SQLQuery("update card set update_status=2 where card_nr = '" & Star_No & "'")
                Else
                    Call SQLQuery("update card set update_status=1 where card_nr = '" & Star_No & "'")
                End If
            End If
            If Mid(txtcard_no, 1, 5) = "CM999" And Linked Then
            If Not IsNumeric(Star_Ext1) Then
                'MsgBox "Format potongan karyawan salah " & vbNewLine & _
                "Harap hubungi IT. ", vbCritical + vbOKOnly, "Oops.."
                'Call isi_data("CM000-00000", "ONE TIME CUSTOMER", 0, "100000", "", "", "")
                'Star_Id = "100000"
                'CmdOk.SetFocus
                'Exit Sub
            Else
            If Star_Ext1 > 10 Then
            MsgBox ("Sisa potongan karyawan anda senilai Rp." + FormatNumber(Star_Ext1, 0, True, True, True) & " !!!")
            End If
            End If
            End If
        End If
        txtcard_opt.SetFocus
    Case 27
        Unload Me
    End Select
End Sub
