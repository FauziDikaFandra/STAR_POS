VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmNumCard 
   Caption         =   "VARIABLE"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNum 
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   18
      Left            =   4050
      TabIndex        =   21
      Top             =   2925
      Width           =   975
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   4050
      TabIndex        =   20
      Top             =   1950
      Width           =   975
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "OM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   16
      Left            =   3075
      TabIndex        =   19
      Top             =   1950
      Width           =   975
   End
   Begin VB.CommandButton btnNum 
      Caption         =   "GM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   4050
      TabIndex        =   18
      Top             =   975
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   5190
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   3150
         Width           =   1950
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "CM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   3075
         TabIndex        =   6
         Top             =   225
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
         Height          =   975
         Index           =   12
         Left            =   3075
         TabIndex        =   5
         Top             =   3150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "SM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   13
         Left            =   3075
         TabIndex        =   4
         Top             =   2175
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
         Index           =   14
         Left            =   4050
         TabIndex        =   3
         Top             =   3150
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5190
      Begin TDBText6Ctl.TDBText txtno 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   225
         TabIndex        =   1
         Top             =   150
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   688
         Caption         =   "Form1.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form1.frx":0066
         Key             =   "Form1.frx":0084
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
         Format          =   "A9"
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
   End
End
Attribute VB_Name = "frmNumCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNum_Click(Index As Integer)
    If Me.Caption = "NUMBER - VIEW" And Index = 10 Then Exit Sub
    
    If Index < 11 Then txtno.Text = txtno.Text + btnNum(Index).Caption
    
    Select Case Index
    Case 11 'Enter
        If Me.Caption = "NUMBER - VIEW" Then
            frmView.txtkode = txtno.Text
            Unload Me
            DoEvents
            frmView.txtkode.SetFocus
            SendKeys "{enter}"
            Exit Sub
        End If
        
        If Me.Caption = "NUMBER - HARGA" Then
            frmHarga.txtprice = txtno.Text
            Unload Me
            DoEvents
            frmHarga.txtprice.SetFocus
            'endKeys "{enter}"
            Exit Sub
        End If
        
        If Me.Caption = "NUMBER - ID" Then
            frmkaryawan.txtid = txtno.Text
            Unload Me
            DoEvents
            frmkaryawan.txtid.SetFocus
            'endKeys "{enter}"
            Exit Sub
        End If
        
        If Me.Caption = "NUMBER - SALES" Then
            frmCard.txtcard_opt = txtno.Text
            Unload Me
            DoEvents
            frmCard.txtcard_opt.SetFocus
            Exit Sub
        End If
    Case 12 'Backspace
        txtno.SetFocus
        SendKeys "{end}+{backspace}"
    Case 13 'Clear
        txtno.Text = txtno.Text + btnNum(Index).Caption
    Case 15 'Clear
        txtno.Text = txtno.Text + btnNum(Index).Caption
    Case 16 'Clear
        txtno.Text = txtno.Text + btnNum(Index).Caption
    Case 17 'Clear
        txtno.Text = txtno.Text + btnNum(Index).Caption
    Case 18 'Clear
        txtno.Text = txtno.Text + btnNum(Index).Caption
    Case 14 'Close
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
If Me.Caption = "NUMBER - HARGA" Then btnNum(10).Enabled = False
End Sub

Private Sub txtno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        btnNum_Click (11)
    Case 27
        Unload Me
    End Select
End Sub


