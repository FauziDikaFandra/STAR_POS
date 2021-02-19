VERSION 5.00
Begin VB.Form frmCicilan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CICILAN"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4050
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
      Height          =   730
      Left            =   2520
      Picture         =   "frmCicilan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
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
      Height          =   730
      Left            =   2520
      Picture         =   "frmCicilan.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "PILIH JENIS CICILAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1750
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton opt12 
         Caption         =   "12 BULAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton opt6 
         Caption         =   "6 BULAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton opt3 
         Caption         =   "3 BULAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCicilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcancel_Click()
    frmPayment.intCicilan = 1
    Unload Me
End Sub

Private Sub Cmdok_Click()
If opt3.Value = True Then
    frmPayment.intCicilan = 3
ElseIf opt6.Value = True Then
    frmPayment.intCicilan = 6
Else
    frmPayment.intCicilan = 12
End If
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        If opt3.Value = True Then
            frmPayment.intCicilan = 3
        ElseIf opt6.Value = True Then
            frmPayment.intCicilan = 6
        Else
            frmPayment.intCicilan = 12
        End If
        Unload Me
    Case 27
        frmPayment.intCicilan = 1
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    opt3.Value = True
End Sub

