VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmCashOpen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CASH OPEN"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashOpen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1140
      Left            =   75
      TabIndex        =   16
      Top             =   0
      Width           =   4215
      Begin TDBText6Ctl.TDBText txtangka 
         Height          =   465
         Left            =   225
         TabIndex        =   0
         Top             =   525
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   820
         Caption         =   "frmCashOpen.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCashOpen.frx":0926
         Key             =   "frmCashOpen.frx":0944
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
         MaxLength       =   7
         LengthAsByte    =   0
         Text            =   "250000"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukan Saldo Awal Transaksi"
         Height          =   315
         Left            =   300
         TabIndex        =   17
         Top             =   225
         Width           =   3315
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   75
      TabIndex        =   1
      Top             =   1125
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1725
         Width           =   975
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
         TabIndex        =   2
         Top             =   3150
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCashOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNum_Click(Index As Integer)
    If Index < 10 Then txtangka.Text = txtangka.Text + btnNum(Index).Caption
    
    Select Case Index
    Case 10
        txtangka.SetFocus
        SendKeys "{end}+{backspace}"
    Case 11
        Call txtangka_KeyDown(13, 0)
    Case 12
        txtangka.Text = ""
    Case 13
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    txtangka.SetFocus
    DoEvents
    SendKeys "{home}+{end}"
End Sub

Private Sub txtangka_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        If txtangka = "" Then Exit Sub
        
        Call SQLQuery("insert into cash(Branch_ID, Datetime, Cash_Register_ID, Shift, User_ID, Modal, Cash, " & _
                    "Voucher , Other_Voucher, Credit_Card, Debet_Card, Credit_Sales, Entertainment, " & _
                    "Deposit , Other_Income, Netto, Discount, Tax, [Returns], No_Sale, Cancel) values " & _
                    "('" & VBranch_ID & "',  convert(varchar(10),getdate(),20), '" & VReg_ID & "', '" & VShift & _
                    "', '" & VKasir_ID & "', " & txtangka.Text & ", 0, " & _
                    "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) ")

        VCopen = True
        Call OpenLaci(0)
        Unload Me
    Case 27
        Unload Me
    End Select
End Sub
