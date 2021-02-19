Attribute VB_Name = "Cetak"
Option Explicit

Private Function Kanan(geser As Byte, rupiah As Long) As String
    Kanan = Space(geser - Len(Format(rupiah, "#,##0"))) & Format(rupiah, "#,##0")
End Function

Public Sub CetakStruk(Status As String, No_trans As String)
Dim vqty As Integer, Vsave As Long, vtotal As Long, Vbayar As Long, abc As String
Dim Rst As New ADODB.Recordset, Rsh As New ADODB.Recordset, AdaCash As Byte
Dim RsX As New ADODB.Recordset

    'cek cash terakhir
    RsX.Open "SELECT TOP (1) Seq From Paid " & _
             "WHERE (Transaction_Number = '" & No_trans & "') AND (Payment_Types = '1') " & _
             "ORDER BY Seq DESC", ConnLocal, adOpenStatic, adLockReadOnly
    
    If Not RsX.EOF Then
        AdaCash = RsX!Seq
    Else
        AdaCash = 0
    End If
    RsX.Close: Set RsX = Nothing
    
    Rsh.Open "select card_number, transaction_date, transaction_time, change_amount, Point_Of_Card_Program, pd.seq as urut, pd.*, pt.* from sales_transactions st inner join paid pd on " & _
             "st.transaction_number = pd.transaction_number inner join payment_types pt on pd.payment_types " & _
             "= pt.payment_types where st.transaction_number='" & No_trans & "' order by pd.seq", ConnLocal, adOpenStatic, adLockReadOnly

    OpenPrintingMode
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(97) & Chr(1) 'tengah
    PrintThisLine Chr(27) & Chr(33) & Chr(0) '10cpi
    PrintThisLine "STAR DEPARTMENT STORE"
    PrintThisLine Right(Tulis(10), Len(Tulis(10)))
'    Printthisline Right(Tulis(10), Len(Tulis(10)) - 5)
'    PrintThisLine Chr(27) & Chr(33) & Chr(1) '12cpi
'    Printthisline Trim(Tulis(1)) nama event
'    Printthisline Trim(Tulis(2)) periode event
    PrintThisLine Chr(27) & Chr(97) & Chr(0) 'kiri
    PrintThisLine "No. " & No_trans
    PrintThisLine VShift & "-" & VKasir_ID & "/" & Left(Trim(VKasir_Nm), 14) & "   " & Format(Rsh!Transaction_Date, "dd/mm/yyyy") & " " & Rsh!Transaction_Time
    PrintThisLine ""
        
    If Status <> "SALES" Then
        PrintThisLine Chr(27) & Chr(97) & Chr(1) 'tengah
        PrintThisLine Chr(27) & Chr(33) & Chr(0) '10cpi
        Select Case Status
        Case "REFUND"
            PrintThisLine "REFUND TRANSACTION"
        Case "REPRINT"
            PrintThisLine "R E P R I N T"
        End Select
        PrintThisLine Chr(27) & Chr(33) & Chr(1) '12cpi
        PrintThisLine Chr(27) & Chr(97) & Chr(0) 'kiri
    End If
    
    Rst.Open "select Seq, sd.PLU, Item_Description, Price, Qty, Discount_Percentage, Discount_Amount, " & _
            "ExtraDisc_Pct, ExtraDisc_Amt, net_price, brand from sales_transaction_details sd inner join item_master im " & _
            "on sd.plu=im.plu where transaction_number='" & No_trans & "' order by seq", ConnLocal, adOpenStatic, adLockReadOnly
    
    While Not Rst.EOF
        PrintThisLine Left(Trim(Rst!plu) & " " & Trim(Rst!item_description), 40)
        abc = "  " & Rst!Qty & "x" & Format(Rst!price, "#,##0") & " " & IIf(Rst!Brand = "No Brand", " ", Rst!Brand)
        abc = Left(abc, 30)
        PrintThisLine abc & Space(40 - Len(CStr(abc)) - Len(Format(Rst!Net_Price, "#,##0"))) & Format(Rst!Net_Price, "#,##0")
       
        If Rst!Discount_Percentage <> 0 Then
            PrintThisLine "  Disc. " & Rst!Discount_Percentage & "% = " & Format(Rst!Discount_Amount, "#,##0")
        End If
        
        If Rst!ExtraDisc_pct <> 0 Then
            PrintThisLine "  Extra " & Rst!ExtraDisc_pct & "% = " & Format(Rst!ExtraDisc_amt, "#,##0")
        End If
        
        vqty = vqty + Rst!Qty
        vtotal = vtotal + Rst!Net_Price
        Vsave = Vsave + Rst!Discount_Amount + Rst!ExtraDisc_amt
        Rst.MoveNext
    Wend
    
    PrintThisLine ""
    PrintThisLine "Total   " & Right("   " & vqty, 4) & " item(s)  : Rp. " & Kanan(12, vtotal)
    PrintThisLine ""
    
    While Not Rsh.EOF
        If Trim(Rsh!credit_card_no) <> "" Then
            If Len(Trim(Rsh!credit_card_no)) = 16 Then
                PrintThisLine Left(Rsh!credit_card_no, 7) & "XXXXXXXXX"
            Else
                PrintThisLine Left(Rsh!credit_card_no, 20)
            End If
            If Trim(Rsh!credit_card_name) <> "" Then PrintThisLine Left(Rsh!credit_card_name, 40)
        End If
        
        abc = Left(Rsh!Description & Space(22), 22) & ": Rp. "
        If Rsh!Payment_Types > 30 Then abc = Left(Rsh!credit_card_name & Space(24), 24) & ": Rp. "
        
        If Trim(Rsh!Description) = "CASH" Then
            If Rsh!urut = AdaCash Then
                PrintThisLine "CASH" & Space(18) & ": Rp. " & Kanan(12, Rsh!paid_amount + Rsh!Change_Amount)
            Else
                PrintThisLine "CASH" & Space(18) & ": Rp. " & Kanan(12, Rsh!paid_amount)
            End If
        Else
            PrintThisLine abc & Space(40 - Len(abc) - Len(Format(Rsh!paid_amount, "#,##0"))) & Format(Rsh!paid_amount, "#,##0")
        End If
        Rsh.MoveNext
    Wend
    
    Rsh.MoveFirst
    
    If Status = "REFUND" Then
        If Rsh!Change_Amount < 0 Then PrintThisLine "CHANGE" & Space(16) & ": Rp. " & Kanan(12, Rsh!Change_Amount)
    Else
        If Rsh!Change_Amount > 0 Then PrintThisLine "CHANGE" & Space(16) & ": Rp. " & Kanan(12, Rsh!Change_Amount)
    End If
    
    PrintThisLine ""
        
    If Vsave > 0 Then
        PrintThisLine "YOU SAVE : Rp. " & Format(Vsave, "#,##0")
        PrintThisLine ""
    End If
    
    If Trim(Rsh!card_number) <> "CM000-00000" Then
        Call MySTAR(Rsh!card_number)
        PrintThisLine "No MySTAR Card  : " & Rsh!card_number
        PrintThisLine "Customer Name   : " & Left(Star_Nm, 22)
        PrintThisLine "Get Point       : " & Get_Point(No_trans)
        'Printthisline "Redeem Point    : " & ""
        If Linked And Status <> "REPRINT" Then PrintThisLine "Point Balance   : " & Star_Pt
        PrintThisLine ""
    End If

    PrintThisLine Chr(27) & Chr(97) & Chr(1) & Tulis(11)
    PrintThisLine Tulis(12)
    PrintThisLine Tulis(13) & ", " & Tulis(14)
    PrintThisLine "NPWP/PKP No : " & Tulis(9)
    PrintThisLine "Harga Sudah Termasuk Pajak"
    PrintThisLine ""
    PrintThisLine Tulis(3)
    PrintThisLine Tulis(4)
    PrintThisLine Tulis(5)
    PrintThisLine ""
    PrintThisLine Tulis(7)
    PrintThisLine Tulis(8)
    PrintThisLine "Facebook / Twitter stardeptstore"
    PrintThisLine Chr(27) & Chr(97) & Chr(0)
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
    
    Rst.Close
    Rsh.Close
ClosePrintingMode
End Sub

Public Sub CetakStruk_Promo(Nama As String, No_trans As String, Pesan As String)
    Open PPort For Output As #1
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine "----------------------------------------"
    PrintThisLine "BILL#    : " & No_trans
    PrintThisLine "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    PrintThisLine "----------------------------------------"
    PrintThisLine Pesan
    PrintThisLine "----------------------------------------"
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
End Sub

Public Sub CetakPesan(Status As String, No_trans As String)
    Open PPort For Output As #1
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine Tulis(11)
    PrintThisLine Tulis(10)
    PrintThisLine "BILL#    : " & No_trans
    PrintThisLine "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Select Case Status
    Case "HOLD"
        PrintThisLine "HOLD TRANSACTION"
    Case "CANCEL"
        PrintThisLine "CANCEL TRANSACTION"
    End Select
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
End Sub

Public Sub CetakBegin()
    OpenPrintingMode
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine "POS BEGIN... "
    PrintThisLine "NPWP     : " & Tulis(9)
    PrintThisLine "REGISTER : " & VReg_ID
    PrintThisLine "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    ClosePrintingMode
End Sub

Public Sub CetakValid(No_trans As String, brs1 As String, brs2 As String)
    Open PPort For Output As #1
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine "BILL#    : " & No_trans
    PrintThisLine "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    PrintThisLine brs1
    PrintThisLine brs2
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
End Sub

Public Sub OpenLaci(tipe As Byte)
    Open PPort For Output As #1
    If tipe = 1 Then
        PrintThisLine Chr(27) + Chr(64) 'reset printer
        PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
        PrintThisLine "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
        PrintThisLine "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
        PrintThisLine "OPEN DRAWER"
        PrintThisLine "": PrintThisLine "": PrintThisLine ""
        PrintThisLine "": PrintThisLine "": PrintThisLine ""
    End If
    PrintThisLine Chr(27) & Chr(112) & Chr(0)
    Close #1
End Sub

Public Sub XRead()
Dim RsBayar As New ADODB.Recordset, Rs As New ADODB.Recordset
Dim Jual As Long, diskon As Long, Retur As Long, Batal As Long, Modal As Long, Jumlah As Long
    
    Call OpenLaci(0)
    
    Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Nilai, isnull(SUM(Total_discount),0) AS Potong " & _
             "FROM Sales_Transactions WHERE Status = '00' and substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
             "' and Transaction_Number in (select transaction_number from paid where Paid.Shift = '" & VShift & "') ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Jual = Rs!nilai
    diskon = Rs!potong
    Rs.Close
    
    Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Balik " & _
             "FROM Sales_Transactions WHERE Flag_Return  = '1' and substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
             "' and Transaction_Number in (select transaction_number from paid where Paid.Shift = '" & VShift & "') ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Retur = Rs!balik
    Rs.Close
     
    Rs.Open "SELECT isnull(SUM(Net_Price),0) AS Nilai " & _
             "FROM Sales_Transaction_Details WHERE substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
             "' and Transaction_Number in (select transaction_number from paid where Paid.Shift = '" & VShift & "' and flag_void='1') ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Batal = Rs!nilai
    Rs.Close
    
    Rs.Open "SELECT Modal From Cash WHERE (Branch_ID = '" & VBranch_ID & "') AND (Datetime = '" & _
            Format(GetSrvDate, "YYYY-MM-DD") & "') AND (Shift = " & VShift & ")", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Modal = Rs!Modal
    Rs.Close: Set Rs = Nothing
    
    RsBayar.Open "SELECT Payment_Types.Description, SUM(Paid.Paid_Amount) AS Nilai " & _
            "FROM Paid INNER JOIN Payment_Types ON Paid.Payment_Types = Payment_Types.Payment_Types " & _
            "WHERE (Paid.Shift = '" & VShift & "') and substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
            "' GROUP BY Payment_Types.seq, Payment_Types.Description order by Payment_Types.seq", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    Open PPort For Output As #1
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine "X-Reading Shift : " & VShift & " " & frmMain.lblline
    PrintThisLine "Branch          : " & Tulis(10)
    PrintThisLine "REGISTER        : " & VReg_ID
    PrintThisLine "CASHIER         : " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME            : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    PrintThisLine "----------------------------------------"
    PrintThisLine "Modal               : Rp. " & Kanan(14, Modal)
    
    While Not RsBayar.EOF
        PrintThisLine Left(RsBayar!Description & Space(20), 20) & ": Rp. " & Kanan(14, RsBayar!nilai)
        Jumlah = Jumlah + RsBayar!nilai
        RsBayar.MoveNext
    Wend
    
    PrintThisLine "----------------------------------------"
    PrintThisLine "TOTAL               : Rp. " & Kanan(14, Jumlah)
    PrintThisLine "OVER VOUCHER        : Rp. " & Kanan(14, Jumlah - Jual)
    PrintThisLine "----------------------------------------"
    PrintThisLine "X Reading           : Rp. " & Kanan(14, Jual)
    PrintThisLine ""
    PrintThisLine "Discount            : Rp. " & Kanan(14, diskon)
    PrintThisLine "Return              : Rp. " & Kanan(14, Retur)
    PrintThisLine "Void                : Rp. " & Kanan(14, Batal)
    
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
    
    ConnLocal.Execute "update cash_register set shift='2' WHERE Branch_ID = '" & VBranch_ID & _
                      "' AND cash_register_id='" & VReg_ID & "'"
    
    'update table cash
    RsBayar.Close: Set RsBayar = Nothing
End Sub

Public Sub ZReset()
Dim RsBayar As New ADODB.Recordset, Rs As New ADODB.Recordset, x As Byte
Dim Jual As Long, diskon As Long, Retur As Long, Batal As Long, Modal As Long, Jumlah As Long
    
    Call OpenLaci(0)
    
    Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Nilai, isnull(SUM(Total_discount),0) AS Potong " & _
             "FROM Sales_Transactions WHERE Status = '00' and substring(transaction_number, 9,8)='" & _
             Format(GetSrvDate, "DDMMYYYY") & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Jual = Rs!nilai
    diskon = Rs!potong
    Rs.Close
    
    Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Balik " & _
             "FROM Sales_Transactions WHERE Flag_Return  = '1' and substring(transaction_number, 9,8)='" & _
             Format(GetSrvDate, "DDMMYYYY") & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Retur = Rs!balik
    Rs.Close
     
    Rs.Open "SELECT isnull(SUM(Net_Price),0) AS Nilai " & _
             "FROM Sales_Transaction_Details WHERE substring(transaction_number, 9,8)='" & _
             Format(GetSrvDate, "DDMMYYYY") & "' and flag_void='1' ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Batal = Rs!nilai
    Rs.Close
    
    Rs.Open "SELECT sum(Modal) as modal From Cash WHERE (Branch_ID = '" & VBranch_ID & "') AND (Datetime = '" & _
            Format(GetSrvDate, "YYYY-MM-DD") & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Modal = Rs!Modal
    Rs.Close: Set Rs = Nothing
    
    RsBayar.Open "SELECT Payment_Types.Description, SUM(Paid.Paid_Amount) AS Nilai " & _
            "FROM Paid INNER JOIN Payment_Types ON Paid.Payment_Types = Payment_Types.Payment_Types " & _
            "WHERE substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
            "' GROUP BY Payment_Types.seq, Payment_Types.Description order by Payment_Types.seq", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    For x = 1 To 3
    Jumlah = 0
    
    Open PPort For Output As #1
    PrintThisLine Chr(27) + Chr(64) 'reset printer
    PrintThisLine Chr(27) & Chr(33) & Chr(1) '10cpi
    PrintThisLine "Z-Reset Shift   : " & VShift & " " & frmMain.lblline
    PrintThisLine "Branch          : " & Tulis(10)
    PrintThisLine "REGISTER        : " & VReg_ID
    PrintThisLine "CASHIER         : " & VKasir_ID & "/" & VKasir_Nm
    PrintThisLine "TIME            : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    PrintThisLine "----------------------------------------"
    PrintThisLine "Modal               : Rp. " & Kanan(14, Modal)
    
    If Not (RsBayar.EOF And RsBayar.BOF) Then RsBayar.MoveFirst
    While Not RsBayar.EOF
        PrintThisLine Left(RsBayar!Description & Space(20), 20) & ": Rp. " & Kanan(14, RsBayar!nilai)
        Jumlah = Jumlah + RsBayar!nilai
        RsBayar.MoveNext
    Wend
    
    PrintThisLine "----------------------------------------"
    PrintThisLine "TOTAL               : Rp. " & Kanan(14, Jumlah)
    PrintThisLine "OVER VOUCHER        : Rp. " & Kanan(14, Jumlah - Jual)
    PrintThisLine "----------------------------------------"
    PrintThisLine "Z Reset             : Rp. " & Kanan(14, Jual)
    PrintThisLine ""
    PrintThisLine "Discount            : Rp. " & Kanan(14, diskon)
    PrintThisLine "Return              : Rp. " & Kanan(14, Retur)
    PrintThisLine "Void                : Rp. " & Kanan(14, Batal)
    
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    PrintThisLine "": PrintThisLine "": PrintThisLine ""
    Close #1
    
    Next x
    'Call SQLQuery("update cash_register set shift='1', last_reset_date=reset_date, reset_date=getdate(), zreset_status=1 WHERE Branch_ID = '" & _
                VBranch_ID & "' AND cash_register_id='" & VReg_ID & "'")
    
    'ConnLocal.Execute "update branches set date_yesterday=date_current, date_current=getdate() " & _
                     "WHERE Branch_ID = '" & VBranch_ID & "'"

    ConnLocal.Execute "exec spp_ZresetLocal '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "'"
    ConnLocal.Execute "exec spp_ZresetServer '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "',''"
    ConnLocal.Execute "exec spp_DeleteTrans"
    
    If Linked Then ConnServer.Execute "exec spp_ZresetServer '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "',''"
    
    RsBayar.Close: Set RsBayar = Nothing
End Sub

