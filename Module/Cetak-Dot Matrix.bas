Attribute VB_Name = "Cetak"
'--- CETAK DOT MATRIX --- CETAK DOT MATRIX --- CETAK DOT MATRIX --- CETAK DOT MATRIX --- CETAK DOT MATRIX ---
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

    Open PPort For Output As #1
    
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(97) & Chr(1); 'tengah
    Print #1, Chr(27) & Chr(33) & Chr(0); '10cpi
    Print #1, "STAR DEPARTMENT STORE"
    Print #1, Right(Tulis(10), Len(Tulis(10)))
'    Print #1, Right(Tulis(10), Len(Tulis(10)) - 5)
    Print #1, Chr(27) & Chr(33) & Chr(1); '12cpi
'    Print #1, Trim(Tulis(1)) nama event
'    Print #1, Trim(Tulis(2)) periode event
    Print #1, Chr(27) & Chr(97) & Chr(0) 'kiri
    Print #1, "No. " & No_trans
    Print #1, VShift & "-" & VKasir_ID & "/" & Left(Trim(VKasir_Nm), 14) & "   " & Format(Rsh!Transaction_Date, "dd/mm/yyyy") & " " & Rsh!Transaction_Time
    Print #1, ""
        
    If Status <> "SALES" Then
        Print #1, Chr(27) & Chr(97) & Chr(1); 'tengah
        Print #1, Chr(27) & Chr(33) & Chr(0); '10cpi
        Select Case Status
        Case "REFUND"
            Print #1, "REFUND TRANSACTION"
        Case "REPRINT"
            Print #1, "R E P R I N T"
        End Select
        Print #1, Chr(27) & Chr(33) & Chr(1); '12cpi
        Print #1, Chr(27) & Chr(97) & Chr(0) 'kiri
    End If
    
    Rst.Open "select Seq, sd.PLU, Item_Description, Price, Qty, Discount_Percentage, Discount_Amount, " & _
            "ExtraDisc_Pct, ExtraDisc_Amt, net_price, brand from sales_transaction_details sd inner join item_master im " & _
            "on sd.plu=im.plu where transaction_number='" & No_trans & "' order by seq", ConnLocal, adOpenStatic, adLockReadOnly
    
    While Not Rst.EOF
        Print #1, Left(Trim(Rst!plu) & " " & Trim(Rst!item_description), 40)
        abc = "  " & Rst!Qty & "x" & Format(Rst!price, "#,##0") & " " & IIf(Rst!Brand = "No Brand", " ", Rst!Brand)
        abc = Left(abc, 30)
        Print #1, abc & Space(40 - Len(CStr(abc)) - Len(Format(Rst!Net_Price, "#,##0"))) & Format(Rst!Net_Price, "#,##0")
       
        If Rst!Discount_Percentage <> 0 Then
            Print #1, "  Disc. " & Rst!Discount_Percentage & "% = " & Format(Rst!Discount_Amount, "#,##0")
        End If
        
        If Rst!ExtraDisc_pct <> 0 Then
            Print #1, "  Extra " & Rst!ExtraDisc_pct & "% = " & Format(Rst!ExtraDisc_amt, "#,##0")
        End If
        
        vqty = vqty + Rst!Qty
        vtotal = vtotal + Rst!Net_Price
        Vsave = Vsave + Rst!Discount_Amount + Rst!ExtraDisc_amt
        Rst.MoveNext
    Wend
    
    Print #1, ""
    Print #1, "Total   " & Right("   " & vqty, 4) & " item(s)  : Rp. " & Kanan(12, vtotal)
    Print #1, ""
    
    While Not Rsh.EOF
        If Trim(Rsh!credit_card_no) <> "" Then
            If Len(Trim(Rsh!credit_card_no)) = 16 Then
                Print #1, Left(Rsh!credit_card_no, 7) & "XXXXXXXXX"
            Else
                Print #1, Left(Rsh!credit_card_no, 20)
            End If
            If Trim(Rsh!credit_card_name) <> "" Then Print #1, Left(Rsh!credit_card_name, 40)
        End If
        
        abc = Left(Rsh!Description & Space(22), 22) & ": Rp. "
        If Rsh!Payment_Types > 30 Then abc = Left(Rsh!credit_card_name & Space(24), 24) & ": Rp. "
        
        If Trim(Rsh!Description) = "CASH" Then
            If Rsh!urut = AdaCash Then
                Print #1, "CASH" & Space(18) & ": Rp. " & Kanan(12, Rsh!paid_amount + Rsh!Change_Amount)
            Else
                Print #1, "CASH" & Space(18) & ": Rp. " & Kanan(12, Rsh!paid_amount)
            End If
        Else
            Print #1, abc & Space(40 - Len(abc) - Len(Format(Rsh!paid_amount, "#,##0"))) & Format(Rsh!paid_amount, "#,##0")
        End If
        Rsh.MoveNext
    Wend
    
    Rsh.MoveFirst
    
    If Status = "REFUND" Then
        If Rsh!Change_Amount < 0 Then Print #1, "CHANGE" & Space(16) & ": Rp. " & Kanan(12, Rsh!Change_Amount)
    Else
        If Rsh!Change_Amount > 0 Then Print #1, "CHANGE" & Space(16) & ": Rp. " & Kanan(12, Rsh!Change_Amount)
    End If
    
    Print #1, ""
        
    If Vsave > 0 Then
        Print #1, "YOU SAVE : Rp. " & Format(Vsave, "#,##0")
        Print #1, ""
    End If
    
    If Trim(Rsh!card_number) <> "CM000-00000" Then
        Call MySTAR(Rsh!card_number)
        Print #1, "No MySTAR Card  : " & Rsh!card_number
        Print #1, "Customer Name   : " & Left(Star_Nm, 22)
        Print #1, "Get Point       : " & Get_Point(No_trans)
        'Print #1, "Redeem Point    : " & ""
        If Linked And Status <> "REPRINT" Then Print #1, "Point Balance   : " & Star_Pt
        Print #1, ""
    End If

    Print #1, Chr(27) & Chr(97) & Chr(1) & Tulis(11)
    Print #1, Tulis(12)
    Print #1, Tulis(13) & ", " & Tulis(14)
    Print #1, "NPWP/PKP No : " & Tulis(9)
    Print #1, "Harga Sudah Termasuk Pajak"
    Print #1, ""
    Print #1, Tulis(3)
    Print #1, Tulis(4)
    Print #1, Tulis(5)
    Print #1, ""
    Print #1, Tulis(7)
    Print #1, Tulis(8)
    Print #1, "Facebook / Twitter stardeptstore"
    Print #1, Chr(27) & Chr(97) & Chr(0)
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
    Close #1
    
    Rst.Close
    Rsh.Close
End Sub

Public Sub CetakStruk_Promo(Nama As String, No_trans As String, Pesan As String)
    Open PPort For Output As #1
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, "----------------------------------------"
    Print #1, "TRANS#    : " & No_trans
    Print #1, "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Print #1, "----------------------------------------"
    Print #1, Pesan
    Print #1, "----------------------------------------"
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
    Close #1
End Sub

Public Sub CetakPesan(Status As String, No_trans As String)
    Open PPort For Output As #1
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, Tulis(11)
    Print #1, Tulis(10)
    Print #1, "TRANS#    : " & No_trans
    Print #1, "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Select Case Status
    Case "HOLD"
        Print #1, "HOLD TRANSACTION"
    Case "CANCEL"
        Print #1, "CANCEL TRANSACTION"
    End Select
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
    Close #1
End Sub

Public Sub CetakBegin()
    Open PPort For Output As #1
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, "POS BEGIN... "
    Print #1, "NPWP     : " & Tulis(9)
    Print #1, "REGISTER : " & VReg_ID
    Print #1, "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
    Close #1
End Sub

Public Sub CetakValid(No_trans As String, brs1 As String, brs2 As String)
    Open PPort For Output As #1
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, "TRANS#    : " & No_trans
    Print #1, "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Print #1, brs1
    Print #1, brs2
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
    Close #1
End Sub

Public Sub OpenLaci(tipe As Byte)
    Open PPort For Output As #1
    If tipe = 1 Then
        Print #1, Chr(27) + Chr(64); 'reset printer
        Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
        Print #1, "CASHIER  : " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
        Print #1, "TIME     : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
        Print #1, "OPEN DRAWER"
        Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
    End If
    Print #1, Chr(27) & Chr(112) & Chr(0)
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
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, "X-Reading Shift : " & VShift & " " & frmMain.lblline
    Print #1, "Branch          : " & Tulis(10)
    Print #1, "REGISTER        : " & VReg_ID
    Print #1, "CASHIER         : " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME            : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Print #1, "----------------------------------------"
    Print #1, "MODAL               : Rp. " & Kanan(14, Modal)
    
    While Not RsBayar.EOF
        Print #1, Left(RsBayar!Description & Space(20), 20) & ": Rp. " & Kanan(14, RsBayar!nilai)
        Jumlah = Jumlah + RsBayar!nilai
        RsBayar.MoveNext
    Wend
    
    Print #1, "----------------------------------------"
    Print #1, "TOTAL               : Rp. " & Kanan(14, Jumlah)
    Print #1, "OVER VOUCHER        : Rp. " & Kanan(14, Jumlah - Jual)
    Print #1, "----------------------------------------"
    Print #1, "X Reading           : Rp. " & Kanan(14, Jual)
    Print #1, ""
    Print #1, "Discount            : Rp. " & Kanan(14, diskon)
    Print #1, "Return              : Rp. " & Kanan(14, Retur)
    Print #1, "Void                : Rp. " & Kanan(14, Batal)
    
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
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
    Print #1, Chr(27) + Chr(64); 'reset printer
    Print #1, Chr(27) & Chr(33) & Chr(1); '10cpi
    Print #1, "Z-Reset Shift   : " & VShift & " " & frmMain.lblline
    Print #1, "Branch          : " & Tulis(10)
    Print #1, "REGISTER        : " & VReg_ID
    Print #1, "CASHIER         : " & VKasir_ID & "/" & VKasir_Nm
    Print #1, "TIME            : " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Print #1, "----------------------------------------"
    Print #1, "MODAL               : Rp. " & Kanan(14, Modal)
    
    If Not (RsBayar.EOF And RsBayar.BOF) Then RsBayar.MoveFirst
    While Not RsBayar.EOF
        Print #1, Left(RsBayar!Description & Space(20), 20) & ": Rp. " & Kanan(14, RsBayar!nilai)
        Jumlah = Jumlah + RsBayar!nilai
        RsBayar.MoveNext
    Wend
    
    Print #1, "----------------------------------------"
    Print #1, "TOTAL               : Rp. " & Kanan(14, Jumlah)
    Print #1, "OVER VOUCHER        : Rp. " & Kanan(14, Jumlah - Jual)
    Print #1, "----------------------------------------"
    Print #1, "Z Reset             : Rp. " & Kanan(14, Jual)
    Print #1, ""
    Print #1, "Discount            : Rp. " & Kanan(14, diskon)
    Print #1, "Return              : Rp. " & Kanan(14, Retur)
    Print #1, "Void                : Rp. " & Kanan(14, Batal)
    
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, ""
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

