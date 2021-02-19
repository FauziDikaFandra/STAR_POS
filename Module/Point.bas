Attribute VB_Name = "Point"
Option Explicit

Public Function MySTAR(No_Kartu As String, no_tlp As String)
Dim RsMySTAR As New ADODB.Recordset
        
    'StrSQL = "select cust_name, isnull(card_point,0) as card_point, cust_id, isnull(cust_address_street2,0) as cust_address_street2, isnull(ext2,0) as ext2, isnull(ext1,0) as ext1, cust_phone_gsm, cust_email, a.update_status from card a inner join customer_master_member b " & _
    '         "on a.Cust_Nr = b.Cust_Nr where Card_Nr = '" & no_kartu & "' and card_status='A'"
If no_tlp = "0" Then
    StrSQL = "select a.card_nr,cust_name, isnull(card_point,0) as card_point, cust_id, " & _
    "isnull(cust_address_street2,0) as cust_address_street2, isnull(ext2,0) as ext2, " & _
    "isnull(ext1,0) as ext1, cust_phone_gsm, cust_email, a.update_status," & _
    "ISNULL(c.Exp_Point_Period,'') aS Exp_Point,ISNULL(c.Expired_Date,'') " & _
    "AS Expired_Date,b.Cust_Gender from card a inner join customer_master_member b on " & _
    "a.Cust_Nr = b.Cust_Nr left join List_Customer_Master_Member c on a.Card_Nr " & _
    "= c.Card_Nr  where a.Card_Nr = '" & No_Kartu & "' and card_status='A'"
Else
    StrSQL = "select a.card_nr,cust_name, isnull(card_point,0) as card_point, cust_id, " & _
    "isnull(cust_address_street2,0) as cust_address_street2, isnull(ext2,0) as ext2, " & _
    "isnull(ext1,0) as ext1, cust_phone_gsm, cust_email, a.update_status," & _
    "ISNULL(c.Exp_Point_Period,'') aS Exp_Point,ISNULL(c.Expired_Date,'') " & _
    "AS Expired_Date,b.Cust_Gender from card a inner join customer_master_member b on " & _
    "a.Cust_Nr = b.Cust_Nr left join List_Customer_Master_Member c on a.Card_Nr " & _
    "= c.Card_Nr  where b.cust_phone_gsm = '" & no_tlp & "' and card_status='A'"
End If

    
    
    If Linked Then
        RsMySTAR.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsMySTAR.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
        
    If Not RsMySTAR.EOF Then
        If no_tlp <> "0" Then
            frmCard.txtcard_no.Text = RsMySTAR!Card_Nr
        End If
        Star_Pt = RsMySTAR!card_point
        Star_Nm = RsMySTAR!cust_name
        Star_Id = RsMySTAR!Cust_id
        Star_Freq = CStr(RsMySTAR!ext2)
        Star_Ext1 = RsMySTAR!ext1
        Star_Omz = CStr(RsMySTAR!cust_address_street2)
        Star_Phone = RsMySTAR!cust_phone_gsm
        Star_Email = RsMySTAR!cust_email
        Star_updsts = RsMySTAR!update_status
        Exp_Point = RsMySTAR!Exp_Point
        Expired_Date = RsMySTAR!Expired_Date
        Star_Gender = RsMySTAR!Cust_Gender
    Else
        Star_Pt = 0
        Star_Nm = "ONE TIME CUSTOMER"
        Star_Id = "100000"
        Star_Freq = ""
        Star_Omz = ""
        Star_Phone = ""
        Star_Email = ""
        Star_updsts = 9
        Exp_Point = ""
        Expired_Date = ""
        Star_Gender = ""
        MsgBox "No Kartu tidak terdaftar / expired" & vbNewLine & "Mohon hubungi information counter", vbOKOnly + vbInformation, "Oops.."
    End If
    RsMySTAR.Close: Set RsMySTAR = Nothing
End Function

Public Sub Save_Point(NoTrans As String, NoCard As String)
On Error GoTo ErrH
Dim zz As Integer, urut As String

    zz = Get_Point(NoTrans)
    urut = Gen_Kode("TM", Right(NoTrans, 4))
    '--perubahan untuk ver cepat
    'ConnLocal.Execute "update card set card_point=card_point+" & zz & " where card_nr = '" & NoCard & "'"
    'ini source 2424
    Call SQLQuery("update card set card_point=card_point+" & zz & " where card_nr = '" & NoCard & "'")
    '----
    StrSQL = "insert into Customer_Transaction_H_MemberCard(CustTrans_Nr, CustTrans_Date, Card_Nr, CustTrans_TotAmount," & _
             "CustTrans_Point, User_ID, Trans_Time, Data_Status) " & _
             "(select '" & urut & "', transaction_date, card_number, net_price, " & zz & ", cashier_id, transaction_time, '00' " & _
             "from Sales_Transactions where Transaction_Number = '" & NoTrans & "')"
    ConnLocal.Execute StrSQL



'    If Linked Then
'        ConnServer.Execute StrSQL
'        ConnLocal.Execute "Update Customer_Transaction_H_MemberCard set data_status='99' where custtrans_nr = '" & urut & "'"
'    End If
    
    StrSQL = "insert into Customer_Transaction_D_MemberCard(CustTrans_Nr, CustTrans_Struk, CustTrans_Amount, Data_Status)" & _
             "select '" & urut & "', Transaction_Number, net_price, '00'" & _
             "from Sales_Transactions where Transaction_Number = '" & NoTrans & "'"
    ConnLocal.Execute StrSQL
    
'    If Linked Then
'        ConnServer.Execute StrSQL
'        ConnLocal.Execute "Update Customer_Transaction_d_MemberCard set data_status='99' where custtrans_nr = '" & urut & "'"
'    End If
    
    
    StrSQL = ""
    Call MySTAR(NoCard, 0)
    MsgBox "Point bertambah : " & zz & vbNewLine & "Saldo Akhir Point : " & Star_Pt, vbInformation, "Oops.."
    'Shell ("Upload_Point.EXE")
    Exit Sub

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog("Save_Point " & Err.Description & " " & Err.Number)
End Sub

Public Function Get_Point(NoTrans As String) As Integer
Dim RsPoint As New ADODB.Recordset
Dim RsPointBank As New ADODB.Recordset
Dim RsKartuKredit As New ADODB.Recordset
Dim Belanja As Long, Npoin As Long, NVoc As Long, Hari As Byte, Kali_Point As Byte

    Hari = IIf(Weekday(GetSrvDate) = 1, 7, Weekday(GetSrvDate) - 1)

    'Cek Bonus Point dari Bank -- AKTIFKAN KALAU ADA KERJASAMA DENGAN BANK SAJA --
'    RsPointBank.Open "select isnull(point,0) as point, substring(activeday," & Hari & ",1) as act_day " & _
'                     "from cust_param_bonus where jenis_kartu='KK' and status_active='1' and GETDATE() between Start and Finish and substring(activeday," & Hari & ",1)='1'", _
'                     ConnLocal, adOpenForwardOnly, adLockReadOnly
'
'    If Not RsPointBank.EOF Then
'        RsKartuKredit.Open "select * from (select transaction_number, isnull(sum(paid_amount),0) as NilaiKK from paid " & _
'               "where paid.transaction_number='" & NoTrans & "' and left(credit_card_no,6) in " & _
'               "(select CAST(nomor AS varchar(6)) from cc_master where CC_Master='KK') " & _
'               "group by transaction_number) aa Inner Join " & _
'               "(select transaction_number,Net_Price as NilaiSS from Sales_Transactions " & _
'               "where transaction_number='" & NoTrans & "') bb " & _
'               "on aa.Transaction_Number = bb.Transaction_Number ", ConnLocal, adOpenForwardOnly, adLockReadOnly
'
'        If Not RsKartuKredit.EOF Then
'            If RsKartuKredit!NilaiKK = RsKartuKredit!nilaiss Then
'                ConnLocal.Execute "update sales_transactions set Point_Of_Card_Program=" & RsPointBank!Point & " where transaction_number='" & NoTrans & "' "
'            End If
'        End If
'        RsKartuKredit.Close: Set RsKartuKredit = Nothing
'    End If
'    RsPointBank.Close: Set RsPointBank = Nothing
    
    RsPoint.Open "SELECT isnull(Amount,0) as amount, substring(active_day," & Hari & ",1) as act_day FROM Cust_Option WHERE (Card_Type = 'CM') ", ConnLocal, adOpenForwardOnly, adLockReadOnly
        If Not RsPoint.EOF Then
            Npoin = IIf(RsPoint!act_day = "1", RsPoint!Amount, 0)
        Else
            Npoin = 0
        End If
    RsPoint.Close

    RsPoint.Open "select card_number, net_price, Point_Of_Card_Program from sales_transactions where transaction_number = '" & NoTrans & "'", ConnLocal, adOpenKeyset, adLockReadOnly
        Belanja = RsPoint!Net_Price
        Kali_Point = RsPoint!Point_Of_Card_Program
        If Left(RsPoint!card_number, 5) = "CM999" Then Npoin = 0
    RsPoint.Close
    
    RsPoint.Open "select isnull(SUM(net_price),0) as rvoc from sales_transaction_details sd inner join item_master im on sd.PLU=im.PLU " & _
                 "where Burui ='NMD92ZZZ9' and Transaction_Number ='" & NoTrans & "'", ConnLocal, adOpenKeyset, adLockReadOnly
        NVoc = RsPoint!RVoc
    RsPoint.Close: Set RsPoint = Nothing
    
    If Npoin > 0 Then
        Get_Point = roundDown((Belanja - NVoc) / Npoin) * Kali_Point
    Else
        Get_Point = 0
    End If
End Function

Public Function Pay_Point(JmlPoint As Integer, NoCard As String, NoTrx As String, rupiah As Long) As String
On Error GoTo ErrH
Dim urut As String

    urut = Gen_Kode("TW", Right(NoTrx, 4))
    
    'ConnLocal.BeginTrans
    'ConnServer.BeginTrans
    
    ConnLocal.Execute "insert into Cust_Point_Trans(Transaction_number, trans_nr, card_nr, current_point, Claim_Point, Claim_Rp, Date_Trans, User_ID, " & _
            "Data_Status) values ('" & NoTrx & "', '" & urut & "', '" & NoCard & "', " & Star_Pt & ", " & JmlPoint & ", " & rupiah & ", getdate(), '" & VKasir_ID & "', '99') "
    
    ConnServer.Execute "insert into Cust_Point_Trans(Transaction_number, trans_nr, card_nr, current_point, Claim_Point, Claim_Rp, Date_Trans, User_ID, " & _
            "Data_Status) values ('" & NoTrx & "', '" & urut & "', '" & NoCard & "', " & Star_Pt & ", " & JmlPoint & ", " & rupiah & ", getdate(), '" & VKasir_ID & "', '99') "
    
    
    'perubahan ipos cepat 18062019
    'ConnLocal.Execute ("Update card set card_point=card_point - " & JmlPoint & "where card_nr = '" & NoCard & "'")
    'ConnServer.Execute ("Update card set card_point=card_point - " & JmlPoint & "where card_nr = '" & NoCard & "'")
    Call SQLQuery("Update card set card_point=card_point - " & JmlPoint & " where card_nr = '" & NoCard & "'")
    
    StrSQL = ""
    
    'ConnLocal.CommitTrans
    'ConnServer.CommitTrans
    
    Pay_Point = urut
    Call MySTAR(NoCard, 0)
    Exit Function
    
ErrH:
    ConnLocal.RollbackTrans
    ConnServer.RollbackTrans
    Pay_Point = "GAGAL"
End Function

Private Function Gen_Kode2(kode As String) As String
Dim RsCari As New ADODB.Recordset
Dim Depan As String

    Select Case kode
    Case "TM"
        Depan = "TM" & Right(VBranch_ID, 3) & VReg_ID & Format(GetSrvDate, "YYMMDD")
        
        StrSQL = "SELECT  max (CAST(RIGHT(custtrans_nr, 4) AS int)) AS nomor " & _
        "FROM Customer_Transaction_H_MemberCard where left(custtrans_nr,14)='" & Depan & "'"
    Case "TW"
        Depan = "TW" & Right(VBranch_ID, 3) & VReg_ID & Format(GetSrvDate, "YYYYMMDD")
        
        StrSQL = "SELECT  max (CAST(RIGHT(trans_nr, 4) AS int)) AS nomor " & _
        "FROM Cust_Point_Trans where left(trans_nr,16)='" & Depan & "'"
    End Select
    
    If Linked Then
        RsCari.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsCari.Open StrSQL, ConnLocal, adOpenStatic, adLockReadOnly
    End If

    If IsNull(RsCari!nomor) Then
        Gen_Kode2 = Depan + "0001"
    Else
        Gen_Kode2 = Depan + Right("000" + CStr(RsCari!nomor + 1), 4)
    End If
    
    RsCari.Close:   Set RsCari = Nothing
End Function

Private Function Gen_Kode(kode As String, belakang As String) As String
Dim RsCari As New ADODB.Recordset
Dim Depan As String

    Select Case kode
    Case "TM"
        Depan = "TM" & Right(VBranch_ID, 3) & VReg_ID & Format(GetSrvDate, "YYMMDD")
        
'        StrSQL = "SELECT  max (CAST(RIGHT(custtrans_nr, 4) AS int)) AS nomor " & _
'        "FROM Customer_Transaction_H_MemberCard where left(custtrans_nr,14)='" & Depan & "'"
    Case "TW"
        Depan = "TW" & Right(VBranch_ID, 3) & VReg_ID & Format(GetSrvDate, "YYYYMMDD")
        
'        StrSQL = "SELECT  max (CAST(RIGHT(trans_nr, 4) AS int)) AS nomor " & _
'        "FROM Cust_Point_Trans where left(trans_nr,16)='" & Depan & "'"
    End Select
    
'    If Linked Then
'        RsCari.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
'    Else
'        RsCari.Open StrSQL, ConnLocal, adOpenStatic, adLockReadOnly
'    End If

'    If IsNull(RsCari!nomor) Then
'        Gen_Kode = Depan + "0001"
'    Else
'        Gen_Kode = Depan + Right("000" + CStr(RsCari!nomor + 1), 4)
'    End If
    
'    RsCari.Close:   Set RsCari = Nothing

    Gen_Kode = Depan + belakang
End Function

Public Function roundDown(dblValue As Double) As Double
On Error GoTo ErrH
Dim myDec As Long
 
    myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
    If myDec > 0 Then
        roundDown = CDbl(Left(CStr(dblValue), myDec))
    Else
        roundDown = dblValue
    End If
    Exit Function
    
ErrH:
    MsgBox Err.Description, vbInformation, "Round Down"
End Function
 
'Public Function roundUp(dblValue As Double) As Double
'On Error GoTo ErrH
'Dim myDec As Long
'
'    myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
'    If myDec > 0 Then
'        roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
'    Else
'        roundUp = dblValue
'    End If
'    Exit Function
'
'ErrH:
'    MsgBox Err.Description, vbInformation, "Round Up"
'End Function
