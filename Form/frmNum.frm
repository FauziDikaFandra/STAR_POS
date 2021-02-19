VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmNum 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VARIABLE"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmNum.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   75
      TabIndex        =   16
      Top             =   75
      Width           =   4215
      Begin TDBText6Ctl.TDBText txtno 
         Height          =   390
         Left            =   225
         TabIndex        =   0
         Top             =   150
         Width           =   3840
         _Version        =   65536
         _ExtentX        =   6773
         _ExtentY        =   688
         Caption         =   "frmNum.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmNum.frx":0930
         Key             =   "frmNum.frx":094E
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
         Format          =   "9A@"
         FormatMode      =   0
         AutoConvert     =   -1
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   75
      TabIndex        =   15
      Top             =   825
      Width           =   4215
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
         Left            =   3075
         TabIndex        =   17
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
         Height          =   975
         Index           =   13
         Left            =   3075
         TabIndex        =   14
         Top             =   2175
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
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "*"
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
         Index           =   10
         Left            =   3075
         TabIndex        =   12
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   225
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmNum"
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
        
        If Me.Caption = "BIN - VIEW" Then
            frmCekBin.txtkode = txtno.Text
            Unload Me
            DoEvents
            frmCekBin.txtkode.SetFocus
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
            frmSales.txtkode = txtno.Text
            Unload Me
            DoEvents
            frmSales.txtkode.SetFocus
            If Right(frmSales.txtkode, 1) = "*" Then
                SendKeys "{end}"
            Else
                SendKeys "{enter}"
            End If
            Exit Sub
        End If
        
        If Me.Caption = "NUMBER - TELP" Then
            If Left(txtno.Text, 1) = "0" Then
                frmCard.txtno_telp = "62" & Mid(txtno.Text, 2, Len(txtno.Text))
            Else
                frmCard.txtno_telp = txtno.Text
            End If
            
            Unload Me
            DoEvents
            frmCard.txtno_telp.SetFocus
            SendKeys "{enter}"
            Exit Sub
        End If
        
        VNomor = VBranch_ID + VReg_ID + "-" + Format(GetSrvDate, "DDMMYYYY") + "-" + Right("000" + CStr(txtno.Text), 4)

        Dim RsCari As New ADODB.Recordset
            RsCari.Open "select status, flag_return,Payment_Program_ID from sales_transactions where transaction_number ='" & _
                        VNomor & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
        
            If RsCari.EOF Then
                MsgBox "Nomor transaksi tidak valid", vbCritical + vbOKOnly, "Oops.."
                GoTo xx
            End If
        
        Select Case Me.Caption
        Case "REPRINT"
            If RsCari!Status = "00" Then
                Call CetakStruk("REPRINT", VNomor)
                Call SaveLog("Reprint Transaction " & VNomor & " " & VSuper_Nm)
                GoTo xx
            End If
        Case "RELEASE"
            If RsCari!Status = "01" Then
                MSCTlp = True
                If Trim(RsCari!Payment_Program_ID) = "1" Then StrukEmail = True
                frmSales.Caption = IIf(RsCari!flag_return = "1", "REFUND", "SALES")
                Call CDisplay(frmSales.Caption, "TRANSACTION")
                Unload Me
                frmSales.Show
                Unload frmMain
                GoTo xx
            End If
        Case "CANCEL"
            If RsCari!Status = "01" Then
                ConnLocal.Execute "Update sales_transactions set net_price=0, net_amount=0, status='02' where transaction_number='" & VNomor & "'"
                Call CetakPesan("CANCEL", VNomor)
                Call SaveLog("Cancel Transaction " & VNomor & " " & VSuper_Nm)
                GoTo xx
            End If
        End Select
        MsgBox "Nomor transaksi tidak valid", vbCritical + vbOKOnly, "Oops.."

xx:
        RsCari.Close: Set RsCari = Nothing
        Unload Me
        VNomor = ""
    Case 12 'Backspace
        txtno.SetFocus
        SendKeys "{end}+{backspace}"
    Case 13 'Clear
        txtno.Text = ""
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
