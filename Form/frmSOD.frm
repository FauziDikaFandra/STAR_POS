VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSOD 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOD"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   ControlBox      =   0   'False
   Icon            =   "frmSOD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   75
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   5940
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Progressive Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   675
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payment Types"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1125
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VWP Current"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "One Plus One"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   150
         TabIndex        =   13
         Top             =   2025
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bin Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   2100
         TabIndex        =   12
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   2100
         TabIndex        =   11
         Top             =   675
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Attribute"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   2100
         TabIndex        =   10
         Top             =   1125
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   2100
         TabIndex        =   9
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Other User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   2100
         TabIndex        =   8
         Top             =   2025
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cash Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   4050
         TabIndex        =   7
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stamp1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   4050
         TabIndex        =   6
         Top             =   675
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stamp2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   12
         Left            =   4050
         TabIndex        =   5
         Top             =   1125
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cpoint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   13
         Left            =   4050
         TabIndex        =   4
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   14
         Left            =   4050
         TabIndex        =   2
         Top             =   2025
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   2550
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   14
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   75
      TabIndex        =   18
      Top             =   525
      Width           =   5940
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SINKRONISASI DATA KE SERVER ...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   75
         TabIndex        =   19
         Top             =   1200
         Width           =   5940
      End
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DOWNLOAD DATA"
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
      Width           =   6090
   End
End
Attribute VB_Name = "frmSOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dbs As String, Svr As String

Private Sub Form_Activate()
    If Me.Caption = "SOD" Then Proses_SOD
    If Me.Caption = "EOD" Then Proses_EOD
End Sub

Private Sub Form_Load()
    Dbs = Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini")
    Svr = "[" & VSvr & "]"
End Sub

Private Sub Proses_SOD()
On Error GoTo ErrH
Dim b As Byte

    lblmsg = "DOWNLOAD DATA.."
    Frame1.Visible = True
    
    ConnLocal.Execute "exec spp_BackupLocalTable"
    'exec spp_DownLoadcpoint '[192.168.1.206]','POS_SERVER'
    For b = 0 To 14
        Chk(b).Value = 1
        DoEvents
        Select Case b
            Case 0: ConnLocal.Execute "exec spp_DownLoadItemMaster '" & Svr & "','" & Dbs & "',''"
            Case 1: ConnLocal.Execute "exec spp_DownLoadProgressivePrice '" & Svr & "','" & Dbs & "'"
            Case 2: ConnLocal.Execute "exec spp_DownLoadPaymentTypes '" & Svr & "','" & Dbs & "'"
            Case 3: ConnLocal.Execute "exec spp_DownLoadVWPCurrent '" & Svr & "','" & Dbs & "'"
            Case 4: ConnLocal.Execute "exec spp_DownLoadOnepOne '" & Svr & "','" & Dbs & "'"
            Case 5: ConnLocal.Execute "exec spp_DownLoadBinCard '" & Svr & "','" & Dbs & "'"
            Case 6: ConnLocal.Execute "exec spp_DownLoadUsers '" & Svr & "','" & Dbs & "'"
            Case 7: ConnLocal.Execute "exec spp_DownLoadBranchAttributes '" & VBranch_ID & "','" & Svr & "','" & Dbs & "'"
            Case 8: ConnLocal.Execute "exec spp_DownLoadMC '" & Svr & "','" & Dbs & "'"
            Case 9: ConnLocal.Execute "exec spp_DownLoadUserBO '" & Svr & "','" & Dbs & "'"
           Case 10: ConnLocal.Execute "exec spp_DownloadCashRegister '" & VReg_ID & "','" & VBranch_ID & "','" & Svr & "','" & Dbs & "'"
           Case 11: ConnLocal.Execute "exec spp_DownLoadStamp1 '" & Svr & "','" & Dbs & "'"
           Case 12: ConnLocal.Execute "exec spp_DownLoadStamp2 '" & Svr & "','" & Dbs & "'"
           Case 13: ConnLocal.Execute "exec spp_DownLoadCpoint '" & Svr & "','" & Dbs & "'"
           Case 14: ConnLocal.Execute "exec spp_DownLoadOthers '" & Svr & "','" & Dbs & "'"
        End Select
        ProgressBar1.Value = b
    Next b
    
    ConnLocal.Execute "update branches set flag_sod=1 where branch_id='" & VBranch_ID & "'"
    Unload Me
    Exit Sub

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "SOD " & Err.Description & " " & Err.Number)
    End
End Sub

Private Sub Proses_EOD()
'On Error GoTo ErrH

    lblmsg = "UPLOAD DATA.."
    Frame1.Visible = False
    DoEvents
    If VPing = "ONLINE" Then
        Call Naikin_Data
        Call Naikin_Promo
        Call Naikin_Point
        
        Dim Fso As New FileSystemObject
        Dim fil As File

        For Each fil In Fso.GetFolder(PathEmail).Files
            Debug.Print
            Dim TabFile() As String
            TabFile = Split(fil.name, ".")
            Kill PathEmail & "\" & fil.name
        Next
        
        'For Each fil In Fso.GetFolder(PathEmail & "\BACKUP").Files
        '    Debug.Print
        '    Dim TabFile2() As String
        '    TabFile2 = Split(fil.name, ".")
        '    Kill PathEmail & "\BACKUP" & "\" & fil.name
        'Next
    End If
'    StrSQL = "exec spp_UnloadData_Handler '" & Svr & "', '" & Dbs & "', '" & VShift & "'"
'    StrSQL = "exec spp_UploadDataToServer '" & Rs!transaction_number & "', '" & Svr & "', '" & Dbs & "', '" & VShift & "'"
'    ConnLocal.Execute StrSQL
    
    Unload Me
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "EOD " & Err.Description & " " & Err.Number)
    End
End Sub

Private Sub Naikin_Data()
Dim RsA As New ADODB.Recordset
    RsA.Open "SELECT Transaction_Number From SALES_TRANSACTIONS Where upload_Status ='00'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If Not RsA.EOF Then
        While Not RsA.EOF
            'ConnLocal.BeginTrans
            StrSQL = "DELETE " & Svr & "." & Dbs & ".dbo.Sales_Transactions  WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.Sales_Transactions " & _
            "(Transaction_Number, Cashier_ID, Customer_ID, Card_Number, Spending_Program_ID, Transaction_Date " & _
            ", Total_Discount, Points_Of_Spending_Program, Point_Of_Item_Program, Point_Of_Card_Program, Payment_program_ID " & _
            ", Branch_ID, Cash_Register_ID, Total_Paid, Net_Price, Tax, Net_Amount, Change_Amount, Flag_Arrange, WorkManShip " & _
            ", Flag_Return, Register_Return, Transaction_Date_Return, Transaction_Number_Return, Last_Point, Get_Point " & _
            ", Status, upload_status, Transaction_Time, Store_Type ) " & _
            "SELECT Transaction_Number, Cashier_ID, Customer_ID, Card_Number, Spending_Program_ID, Transaction_Date, " & _
            "Total_Discount, Points_Of_Spending_Program, Point_Of_Item_Program, Point_Of_Card_Program, Payment_program_ID, " & _
            "Branch_ID, Cash_Register_ID, Total_Paid, Net_Price, Tax, Net_Amount, Change_Amount, Flag_Arrange, WorkManShip, " & _
            "Flag_Return, Register_Return, Transaction_Date_Return, Transaction_Number_Return, Last_Point, Get_Point, " & _
            "Status , upload_status, Transaction_Time, Store_Type " & _
            "FROM SALES_TRANSACTIONS  WHERE transaction_number= '" & RsA!Transaction_Number & "'"

            ConnLocal.Execute StrSQL
            
            StrSQL = "DELETE " & Svr & "." & Dbs & ".dbo.Sales_Transaction_Details WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.Sales_Transaction_details " & _
            "(Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, Discount_Percentage,  " & _
            "Discount_Amount, extradisc_pct, extradisc_amt, Net_Price, Points_Received, Flag_Void, Flag_Status, Flag_Paket_Discount) " & _
            "SELECT Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, Discount_Percentage,  " & _
            "Discount_Amount, extradisc_pct, extradisc_amt, Net_Price, Points_Received, Flag_Void, Flag_Status, Flag_Paket_Discount " & _
            "FROM Sales_Transaction_details WHERE transaction_number= '" & RsA!Transaction_Number & "'"

            ConnLocal.Execute StrSQL
            
            StrSQL = "DELETE " & Svr & "." & Dbs & ".dbo.paid WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.paid " & _
            "(Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No,  " & _
            "Credit_Card_Name, Paid_Amount, Shift)  " & _
            "SELECT Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No,   " & _
            "Credit_Card_Name, Paid_Amount, Shift  " & _
            "From PAID  WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            
            ConnLocal.Execute StrSQL
  
            StrSQL = "update sales_transactions set upload_status='99' WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "delete from " & Svr & "." & Dbs & ".dbo.cash where branch_id='" & VBranch_ID & "' and Cash_Register_ID ='" & VReg_ID & "' and Datetime = '" & _
            Format(GetSrvDate, "YYYY-MM-DD") & "'"
            
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert  " & Svr & "." & Dbs & ".dbo.cash " & _
            "(Branch_ID, Datetime, Cash_Register_ID, Shift, User_ID, Modal, Cash, Voucher, " & _
            "Other_Voucher, Credit_Card, Debet_Card, Credit_Sales, Entertainment, Deposit, Other_Income, Netto, " & _
            "Discount, Tax, [Returns] , No_Sale, Cancel) " & _
            "SELECT Branch_ID, Datetime, Cash_Register_ID, Shift, User_ID, Modal, Cash, Voucher, " & _
            "Other_Voucher, Credit_Card, Debet_Card, Credit_Sales, Entertainment, Deposit, Other_Income, Netto,  " & _
            "Discount, Tax, [Returns], No_Sale, Cancel  " & _
            "FROM Cash where branch_id='" & VBranch_ID & "' and Cash_Register_ID ='" & VReg_ID & "' and Datetime = '" & _
            Format(GetSrvDate, "YYYY-MM-DD") & "'"
            
            ConnLocal.Execute StrSQL
            
            'ConnLocal.CommitTrans
            RsA.MoveNext
        Wend
    End If
    
    RsA.Close: Set RsA = Nothing
    Exit Sub

ErrH:
    'ConnLocal.RollbackTrans
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Naikin_Data " & Err.Description & " " & Err.Number)
    End
End Sub

Private Sub Naikin_Promo()
Dim RsA As New ADODB.Recordset
    RsA.Open "SELECT Transaction_Number, qty_promo, promo_id From promo_sales Where Status ='00'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If Not RsA.EOF Then
        While Not RsA.EOF
            StrSQL = "DELETE " & Svr & "." & Dbs & ".dbo.promo_sales WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.promo_sales " & _
            "(promo_id, transaction_number, nilai, qty_promo, status ) " & _
            "SELECT  promo_id, transaction_number, nilai, qty_promo, '99'" & _
            "FROM promo_sales  WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "update promo_sales set status='99' WHERE transaction_number= '" & RsA!Transaction_Number & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "update " & Svr & "." & Dbs & ".dbo.promo_hdr set qtyout=qtyout+ " & RsA!qty_promo & " where promo_id='" & RsA!promo_id & "' and islimit=1"
            ConnLocal.Execute StrSQL
                    
            RsA.MoveNext
        Wend
    End If
    
    RsA.Close: Set RsA = Nothing
    Exit Sub

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Naikin_Promo " & Err.Description & " " & Err.Number)
    End
End Sub

Private Sub Naikin_Point()
Dim RsA As New ADODB.Recordset
    RsA.Open "SELECT CustTrans_Nr, CustTrans_Date, Card_Nr, CustTrans_TotAmount, CustTrans_Point, User_ID, Trans_Time, " & _
             "Data_Status From customer_transaction_h_membercard Where data_Status ='00'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    'update 29/9/2016 ojie
    'RsA.Open "SELECT * From customer_transaction_h_membercard a inner join customer_transaction_d_membercard b " & _
             '"On a.CustTrans_Nr =  b.CustTrans_Nr Where a.data_Status ='00' or b.data_Status = '00'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    If Not RsA.EOF Then
        While Not RsA.EOF
            StrSQL = "DELETE FROM " & Svr & "." & Dbs & ".dbo.customer_transaction_h_membercard WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "DELETE FROM " & Svr & "." & Dbs & ".dbo.customer_transaction_d_membercard WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.customer_transaction_h_membercard " & _
            "(CustTrans_Nr, CustTrans_Date, Card_Nr, CustTrans_TotAmount, CustTrans_Point, User_ID, Trans_Time, Data_Status ) " & _
            "SELECT  CustTrans_Nr, CustTrans_Date, Card_Nr, CustTrans_TotAmount, CustTrans_Point, User_ID, Trans_Time, '00'" & _
            "FROM customer_transaction_h_membercard  WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.customer_transaction_d_membercard " & _
            "(CustTrans_Nr, CustTrans_Struk, CustTrans_Amount, Data_Status) " & _
            "SELECT  CustTrans_Nr, CustTrans_Struk, CustTrans_Amount, '00'" & _
            "FROM customer_transaction_d_membercard  WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "update customer_transaction_h_membercard set data_status='99' WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            StrSQL = "update customer_transaction_d_membercard set data_status='99' WHERE CustTrans_Nr= '" & RsA!CustTrans_Nr & "'"
            ConnLocal.Execute StrSQL
            
            'StrSQL = "update " & Svr & "." & Dbs & ".dbo.card set card_point=card_point+ " & RsA!CustTrans_Point & " where card_nr='" & RsA!Card_Nr & "'"
            'ConnLocal.Execute StrSQL
                    
            'Apabila update card ke server dimatikan pada saat save point source 2424 "kurang efektif"
            'StrSQL = "IF EXISTS (select * from Card a inner join " & Svr & "." & Dbs & ".dbo.Card b on a.Card_Nr = b.Card_Nr where a.Card_Point <> b.Card_Point and a.card_nr = '" & RsA!Card_Nr & "') BEGIN update " & Svr & "." & Dbs & ".dbo.card set card_point=card_point+ " & RsA!CustTrans_Point & " where card_nr='" & RsA!Card_Nr & "' END"
            'ConnLocal.Execute StrSQL
            
            RsA.MoveNext
        Wend
    End If
    
    RsA.Close: Set RsA = Nothing
    Exit Sub

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Naikin_Point " & Err.Description & " " & Err.Number)
    End
End Sub

