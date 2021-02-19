VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmCekBin 
   Caption         =   "Cek No BIN Kartu"
   ClientHeight    =   5685
   ClientLeft      =   3345
   ClientTop       =   690
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8370
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
      Left            =   5635
      Picture         =   "CekBin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
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
      Left            =   4440
      Picture         =   "CekBin.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1140
   End
   Begin VB.CommandButton cmdangka 
      Caption         =   "&NUM"
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
      Left            =   3225
      Picture         =   "CekBin.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   1140
   End
   Begin TDBText6Ctl.TDBText txtkode 
      Height          =   390
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   688
      Caption         =   "CekBin.frx":198E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CekBin.frx":19F0
      Key             =   "CekBin.frx":1A0E
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
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   18
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
   Begin TDBText6Ctl.TDBText txtinfo 
      Height          =   4530
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
      _ExtentY        =   7990
      Caption         =   "CekBin.frx":1A42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CekBin.frx":1AAE
      Key             =   "CekBin.frx":1ACC
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
      MultiLine       =   -1
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
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
End
Attribute VB_Name = "frmCekBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCek As New ADODB.Recordset
Private Sub cmdangka_Click()
    frmNum.Caption = "BIN - VIEW"
    frmNum.Show 1
End Sub

Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub Cmdok_Click()
    If txtkode = "" Then Exit Sub
    RsCek.Open "select promo_id,promo_name + ' ** Mulai Dari Tanggal ' + convert(NVARCHAR, start_date , 106) + ' s/d ' + " & _
                "convert(NVARCHAR, end_date , 106) + ' Min Belanja Non Member Rp.' + REPLACE(CONVERT(varchar(20), (CAST(min_purchase AS money)), 1), '.00', '') + ' dan Min Belanja Member Rp.' + " & _
                "REPLACE(CONVERT(varchar(20), (CAST(min_member AS money)), 1), '.00', '') + ' ' + txt1 + ' ' + txt2 + ' **' as Description   from CC_master a " & _
                "inner join Promo_Hdr b on a.CC_Master = b.promo_id where getdate() between start_date and end_date and Nomor  = '" & Trim(txtkode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    txtinfo = "PROMO YANG BERLAKU UNTUK BIN : " + txtkode + " ADALAH : " + vbNewLine + vbNewLine
    If Not RsCek.EOF Then
        While Not RsCek.EOF
            txtinfo = txtinfo & "*** " & RsCek!promo_id & " ***" & vbNewLine
            txtinfo = txtinfo & RsCek!Description & vbNewLine & vbNewLine
            'txtinfo = UCase(txtinfo)
            RsCek.MoveNext
        Wend
    Else
        txtinfo = "*** BIN DENGAN KODE : " + txtkode + " TIDAK TERDAFTAR DALAM PROMO APAPUN ***"
    End If
    
    
    Me.Caption = "VIEW BIN PROMO - " & RsCek.RecordCount
    RsCek.Close
    txtkode.SetFocus
End Sub


Private Sub Form_Activate()
        txtkode.SetFocus
End Sub


Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
    If txtkode = "" Then Exit Sub
    RsCek.Open "select promo_id,promo_name + ' ** Mulai Dari Tanggal ' + convert(NVARCHAR, start_date , 106) + ' s/d ' + " & _
                "convert(NVARCHAR, end_date , 106) + ' Min Belanja Non Member Rp.' + REPLACE(CONVERT(varchar(20), (CAST(min_purchase AS money)), 1), '.00', '') + ' dan Min Belanja Member Rp.' + " & _
                "REPLACE(CONVERT(varchar(20), (CAST(min_member AS money)), 1), '.00', '') + ' ' + txt1 + ' ' + txt2 + ' **' as Description   from CC_master a " & _
                "inner join Promo_Hdr b on a.CC_Master = b.promo_id where getdate() between start_date and end_date and Nomor  = '" & Trim(txtkode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    txtinfo = "PROMO YANG BERLAKU UNTUK BIN : " + txtkode + " ADALAH : " + vbNewLine + vbNewLine
    If Not RsCek.EOF Then
        While Not RsCek.EOF
            txtinfo = txtinfo & "*** " & RsCek!promo_id & " ***" & vbNewLine
            txtinfo = txtinfo & RsCek!Description & vbNewLine & vbNewLine
            'txtinfo = UCase(txtinfo)
        RsCek.MoveNext
        Wend
    Else
        txtinfo = "*** BIN DENGAN KODE : " + txtkode + " TIDAK TERDAFTAR DALAM PROMO APAPUN ***"
    End If
    
    
    Me.Caption = "VIEW BIN PROMO - " & RsCek.RecordCount
    RsCek.Close
    txtkode.SetFocus
    Case 27
        Unload Me
    End Select
End Sub
