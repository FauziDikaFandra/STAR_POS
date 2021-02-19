VERSION 5.00
Begin VB.Form frmDisc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPTION"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2730
   ControlBox      =   0   'False
   Icon            =   "frmDisc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   2730
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1575
      Picture         =   "frmDisc.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
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
      Left            =   1575
      Picture         =   "frmDisc.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2175
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
      Left            =   1575
      Picture         =   "frmDisc.frx":1CCE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
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
      Left            =   1575
      Picture         =   "frmDisc.frx":26D0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   525
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
      Height          =   3420
      ItemData        =   "frmDisc.frx":30D2
      Left            =   0
      List            =   "frmDisc.frx":30D4
      TabIndex        =   1
      Top             =   450
      Width           =   1590
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DISCOUNT"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2790
   End
End
Attribute VB_Name = "frmDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdup_Click()
    List1.ListIndex = IIf(List1.ListIndex > 0, List1.ListIndex - 1, 0)
End Sub

Private Sub cmddown_Click()
    List1.ListIndex = IIf(List1.ListIndex < List1.ListCount - 1, List1.ListIndex + 1, List1.ListCount - 1)
End Sub

Private Sub Cmdok_Click()
    Select Case lblmsg.Caption
    Case "DISCOUNT"
    
    If frmSales.vdisc1 = "0" Then
        frmSales.vdisc1 = List1.Text
    Else
        frmSales.vdisc2 = List1.Text
    End If
    
    Case "VALIDASI"
        If List1.Text = "TOTAL" Then
            Call CetakValid(VNomor, "Total Rp. " & Format(frmSales.vgtotal, "#,##0"), "")
            Unload Me
            Exit Sub
        End If
        
        Dim RsV As New ADODB.Recordset
        Dim aa As String, bb As String
        Dim cc As Long
        
            RsV.Open "select sd.*,it.Long_Description from sales_transaction_details sd inner join item_master it " & _
                    "on sd.plu=it.plu where transaction_number='" & VNomor & "' and brand='" & UbahChar(List1.Text) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            While Not RsV.EOF
                aa = aa & vbNewLine & RsV!plu & " " & RsV!Qty & " X " & Format(RsV!price, "#,##0") & " = " & "Rp. " & Format(RsV!Net_Price, "#,##0")
                If RsV!Discount_Percentage > 0 Then aa = aa & vbNewLine & "Disc." & RsV!Discount_Percentage & "% = " & Format(RsV!Discount_Amount, "#,##0")
                If RsV!ExtraDisc_pct > 0 Then aa = aa & vbNewLine & "Disc." & RsV!ExtraDisc_pct & "% = " & Format(RsV!ExtraDisc_amt, "#,##0")
                aa = aa & vbNewLine & RsV!Long_Description
                cc = cc + RsV!Net_Price
                RsV.MoveNext
            Wend
            
            RsV.MoveFirst
            bb = vbNewLine & RsV!flag_paket_discount & "/" & Left(Siapa_SPG(RsV!flag_paket_discount), 10)
            bb = vbNewLine & "Total " & List1.Text & " Rp. " & Format(cc, "#,##0") & bb
            Call CetakValid(VNomor, aa, bb)
            Call SaveLog("Validasi Kasir : " & VKasir_ID & "/" & VKasir_Nm & "SPG : " & RsV!flag_paket_discount & "/" & Left(Siapa_SPG(RsV!flag_paket_discount), 10))
            RsV.Close: Set RsV = Nothing
    Case "VOUCHER"
          frmPayment.txtno_voc = List1.Text & "-"
    End Select
    
    Unload Me
End Sub

Private Function Siapa_SPG(kode) As String
Dim RsCari As New ADODB.Recordset
            
    RsCari.Open "select spg_id, spg_name from spg where spg_id = '" & _
                Trim(kode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

    Siapa_SPG = IIf(Not RsCari.EOF, RsCari!spg_name, "")

    RsCari.Close: Set RsCari = Nothing
End Function

Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim RsCari As New ADODB.Recordset
    List1.Clear
    
    Select Case lblmsg.Caption
    Case "DISCOUNT"
        RsCari.Open "select disc_1,disc_2,disc_3,disc_4,disc_5,disc_6,disc_7 from cash_register where branch_id = '" & VBranch_ID & _
                "' and cash_register_id = '" & VReg_ID & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                
        List1.AddItem RsCari!disc_1
        List1.AddItem RsCari!disc_2
        List1.AddItem RsCari!disc_3
        List1.AddItem RsCari!disc_4
        List1.AddItem RsCari!disc_5
        List1.AddItem RsCari!disc_6
        List1.AddItem RsCari!disc_7
        
    Case "VALIDASI"
        RsCari.Open "select distinct brand from sales_transaction_details sd inner join item_master it " & _
                    "on sd.plu=it.plu where transaction_number='" & VNomor & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
        
        List1.AddItem "TOTAL"
        While Not RsCari.EOF
            List1.AddItem RsCari!Brand
            RsCari.MoveNext
        Wend
        
    Case "VOUCHER"
        List1.AddItem "AA"
        List1.AddItem "AF"
        List1.AddItem "AM"
        List1.AddItem "AS"
        List1.AddItem "BG"
        List1.AddItem "BQ"
        List1.AddItem "BR"
        List1.AddItem "BS"
        List1.AddItem "BW"
        List1.AddItem "BY"
        List1.AddItem "CB"
        List1.AddItem "CJ"
        List1.AddItem "CK"
        List1.AddItem "CM"
        List1.AddItem "CN"
        List1.AddItem "CY"
        List1.AddItem "DB"
        List1.AddItem "DE"
        List1.AddItem "DH"
        List1.AddItem "DK"
        List1.AddItem "DN"
        List1.AddItem "DP"
        List1.AddItem "ZA"
        List1.AddItem "ZB"
        List1.AddItem "ZR"
        List1.ListIndex = 0
        Exit Sub
    End Select
    
    RsCari.Close:   Set RsCari = Nothing
    List1.ListIndex = 0
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Cmdok_Click
    Case 27
        Cmdcancel_Click
    End Select
End Sub
