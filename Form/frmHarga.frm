VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmHarga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTER PRICE"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   ControlBox      =   0   'False
   Icon            =   "frmHarga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1425
      Picture         =   "frmHarga.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1125
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
      Left            =   2625
      Picture         =   "frmHarga.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1125
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
      Left            =   225
      Picture         =   "frmHarga.frx":1CCE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1125
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCAN BARCODE HARGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   3840
      Begin TDBText6Ctl.TDBText txtprice 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   688
         Caption         =   "frmHarga.frx":2258
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHarga.frx":22B4
         Key             =   "frmHarga.frx":22D2
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
         MaxLength       =   13
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
   End
End
Attribute VB_Name = "frmHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdangka_Click()
    frmNum.Caption = "NUMBER - HARGA"
    frmNum.Show 1
End Sub

Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub Cmdok_Click()
    txtprice.SetFocus
    SendKeys "{enter}"
End Sub

Private Sub txtprice_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Berapa As Double

    Select Case KeyCode
    Case 13
        If txtprice.Text = "" Then Exit Sub
        Berapa = txtprice.Text
        If Left(txtprice.Text, 2) = "27" And Len(txtprice.Text) = 13 Then
            Berapa = CDbl(Mid(txtprice.Text, 3, 10))
        End If
        If Berapa > 99999998 Then
            MsgBox "Harga yang anda masukkan salah", vbCritical + vbOKOnly, "Oops.."
            DoEvents
            txtprice.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        frmSales.vharga = Berapa
        VOK = True
        Unload Me
    Case 27
        Unload Me
        frmSales.txtkode = ""
        frmSales.txtkode.SetFocus
    End Select
End Sub

