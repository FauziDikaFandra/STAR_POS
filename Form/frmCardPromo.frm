VERSION 5.00
Begin VB.Form frmCardPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARD PROMO"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   ControlBox      =   0   'False
   Icon            =   "frmCardPromo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
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
      Height          =   840
      Left            =   3375
      Picture         =   "frmCardPromo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   840
      Left            =   2250
      Picture         =   "frmCardPromo.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmddown 
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1125
      Picture         =   "frmCardPromo.frx":1CCE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmdup 
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   0
      Picture         =   "frmCardPromo.frx":26D0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1725
      Width           =   1140
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      ItemData        =   "frmCardPromo.frx":30D2
      Left            =   0
      List            =   "frmCardPromo.frx":30D4
      TabIndex        =   0
      Top             =   450
      Width           =   4515
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "KARTU PROMOSI"
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
      TabIndex        =   3
      Top             =   0
      Width           =   4515
   End
End
Attribute VB_Name = "frmCardPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdup_Click()
On Error Resume Next
    List1.ListIndex = IIf(List1.ListIndex > 0, List1.ListIndex - 1, 0)
End Sub

Private Sub cmddown_Click()
    List1.ListIndex = IIf(List1.ListIndex < List1.ListCount - 1, List1.ListIndex + 1, List1.ListCount - 1)
End Sub

Private Sub Cmdok_Click()
    frmCard.Vpromo_id = Left(List1.Text, 6)
    Unload Me
End Sub

Private Sub Cmdcancel_Click()
    frmCard.Vpromo_id = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim RsCard As New ADODB.Recordset

    RsCard.Open "select card_promo_id + '    ' + card_promo_name + '-' + Card_Promo_Name_Long as id " & _
                "from Card_Promotion_Name where GETDATE() between Start_Promo_Date and End_Promo_Date ", _
                ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If Not RsCard.EOF Then
        While Not RsCard.EOF
            List1.AddItem RsCard!id
            RsCard.MoveNext
        Wend
        List1.ListIndex = 0
    End If
    
    RsCard.Close:   Set RsCard = Nothing
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Cmdok_Click
    Case 27
        Cmdcancel_Click
    End Select
End Sub

