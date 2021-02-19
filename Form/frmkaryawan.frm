VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmkaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTER ID"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "MASUKAN ID KARYAWAN"
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
      TabIndex        =   4
      Top             =   75
      Width           =   3840
      Begin TDBText6Ctl.TDBText txtid 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   3540
         _Version        =   65536
         _ExtentX        =   6244
         _ExtentY        =   688
         Caption         =   "frmkaryawan.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmkaryawan.frx":005C
         Key             =   "frmkaryawan.frx":007A
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
      Picture         =   "frmkaryawan.frx":00CC
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Picture         =   "frmkaryawan.frx":0656
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1125
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
      Left            =   1425
      Picture         =   "frmkaryawan.frx":1058
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1125
      Width           =   1140
   End
End
Attribute VB_Name = "frmkaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdangka_Click()
    frmNum.Caption = "NUMBER - ID"
    frmNum.Show 1
End Sub

Private Sub Cmdcancel_Click()
    VTanya = True
    Unload Me
End Sub

Private Sub Cmdok_Click()
If Len(txtid.Text) = 9 Then
txtid.SetFocus
SendKeys "{enter}"
Else
MsgBox "No Karyawan Tidak Sesuai !!!", vbCritical + vbOKOnly, "Oops.."
Exit Sub
End If
End Sub

Private Sub txtid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        VKary = txtid
        VTanya = True
        Unload Me
        frmSales.txtkode.SetFocus
    Case 27
        VKary = ""
        VTanya = True
        Unload Me
        frmSales.txtkode.SetFocus
    End Select
End Sub


