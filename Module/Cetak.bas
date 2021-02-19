Attribute VB_Name = "Cetak"
'--- CETAK DOT THERMAL --- CETAK DOT THERMAL --- CETAK DOT THERMAL --- CETAK DOT THERMAL --- CETAK DOT THERMAL ---
Option Explicit
Public stat As Boolean
Dim PosY As Variant
Private Function Kanan(geser As Byte, rupiah As Long) As String
    Kanan = Space(geser - Len(Format(rupiah, "#,##0"))) & Format(rupiah, "#,##0")
End Function

Private Function CetakTengah(tex1 As String) As String
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(tex1)) \ 2
    Printer.Print tex1
End Function

Private Function CetakKanan(tex1 As String) As String
    '----------------------Normal-------------------------
    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(tex1) - 90
    '----------------------POSIFLEX-----------------------
    'Printer.CurrentX = Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Tex1) - 270
    Printer.Print tex1
End Function

Private Function CetakKonon(tex1 As String) As String
    Printer.CurrentX = 3600 - Printer.TextWidth(tex1) - 90
    Printer.Print tex1
End Function


Private Sub Turun()
    Printer.CurrentY = Printer.CurrentY + 60
End Sub

Private Function PosYY(cnt As Variant) As Variant
    PosY = PosY + cnt
    PosYY = PosY
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
    
    Printer.Font.name = "Printer 14cpi"
    Printer.Font.Size = 10
    
    
    If Cfg_Get("Device", "Use_Logo", App.Path & "\config.ini") = 0 Then
    CetakTengah ("STAR DEPARTMENT STORE")
    Else
    CetakTengah ("DEPARTMENT STORE")
    End If
    
    
    Turun
'    Printer.Font.Size = 10
'    CetakTengah ("STAR 8th Anniversary")
'    Turun
    Printer.Font.Size = 9
    CetakTengah (Tulis(10))
    
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print ""

    Printer.Print "No. " & No_trans
    Turun
    Printer.Print VShift & "-" & VKasir_ID & "/" & Left(Trim(VKasir_Nm), 14) & "     " & Format(Rsh!Transaction_Date, "dd/mm/yyyy") & " " & Rsh!Transaction_Time
    Printer.Print ""

    If Status <> "SALES" Then

        Select Case Status
        Case "REFUND"
            Printer.Font.Size = 9
            CetakTengah ("REFUND TRANSACTION")
        Case "REPRINT"
            Printer.Font.Size = 9
            CetakTengah ("R E P R I N T")
        End Select
        Printer.Print ""
    End If

    Rst.Open "select Seq, sd.PLU, Item_Description, Price, Qty, Discount_Percentage, Discount_Amount, " & _
            "ExtraDisc_Pct, ExtraDisc_Amt, net_price, brand from sales_transaction_details sd inner join item_master im " & _
            "on sd.plu=im.plu where transaction_number='" & No_trans & "' order by seq", ConnLocal, adOpenStatic, adLockReadOnly
    Printer.Font.Size = 8

    While Not Rst.EOF
        Turun
        Printer.Print Left(Trim(Rst!plu) & " " & Trim(Rst!Item_Description), 42)
        abc = "  " & Rst!Qty & "x" & Format(Rst!price, "#,##0") & " " & IIf(Rst!Brand = "No Brand", " ", Rst!Brand)
        abc = Left(abc, 30)

        Turun
        Printer.Print abc;
        '---------------------Normal----------------
        CetakKanan (Format(Rst!Net_Price, "#,##0"))
        '---------------------POSIFLEX----------------
        'CetakKonon (Format(Rst!Net_Price, "#,##0"))
        If Rst!Discount_Percentage <> 0 Then
            Turun
            Printer.Print "  Disc. " & Rst!Discount_Percentage & "% = " & Format(Rst!Discount_Amount, "#,##0")
        End If

        If Rst!ExtraDisc_pct <> 0 Then
            Turun
            Printer.Print "  Extra " & Rst!ExtraDisc_pct & "% = " & Format(Rst!ExtraDisc_amt, "#,##0")
        End If

        vqty = vqty + Rst!Qty
        vtotal = vtotal + Rst!Net_Price
        Vsave = Vsave + Rst!Discount_Amount + Rst!ExtraDisc_amt
        Rst.MoveNext
    Wend

    Printer.Print ""
    Printer.Print "Total   " & Right("   " & vqty, 4) & " item(s)   ";
    '--------------Normal------------------------
    Printer.CurrentX = 2000: Printer.Print "  : Rp. ";
    '--------------POSIFLEX----------------------
    'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
    CetakKanan (Kanan(14, vtotal))
    Printer.Print ""

    Printer.Font.Size = 8

    While Not Rsh.EOF
        If Trim(Rsh!Credit_Card_No) <> "" Then
            Turun
            If Len(Trim(Rsh!Credit_Card_No)) = 16 Then
                Printer.Print Left(Rsh!Credit_Card_No, 7) & "XXXXXXXXX"
            Else
                Printer.Print Left(Rsh!Credit_Card_No, 20)
            End If
            'If Trim(Rsh!credit_card_name) <> "" Then Printer.Print Left(Rsh!credit_card_name, 40)
        End If

        abc = Left(Rsh!Description & Space(22), 22)
        If Rsh!Payment_Types > 30 Then abc = Left(Rsh!credit_card_name & Space(24), 24)

        Turun
        If Trim(Rsh!Description) = "CASH" Then
            Printer.Print "CASH";
            '--------------Normal------------------------
            Printer.CurrentX = 2000: Printer.Print "  : Rp. ";
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            If Rsh!urut = AdaCash Then
                CetakKanan Format((Rsh!paid_amount + Rsh!Change_Amount), "#,##0")
            Else
                CetakKanan Format(Rsh!paid_amount, "#,##0")
            End If
        Else
            Printer.Print abc;
            '--------------Normal------------------------
            Printer.CurrentX = 2000: Printer.Print "  : Rp. ";
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            CetakKanan (Format(Rsh!paid_amount, "#,##0"))
        End If
        Rsh.MoveNext
    Wend

    Rsh.MoveFirst
    Turun
    If Status = "REFUND" Then
        If Rsh!Change_Amount < 0 Then
            Printer.Print "CHANGE";
            Printer.CurrentX = 2000: Printer.Print "  : Rp. "
            CetakKanan (Format(Rsh!Change_Amount, "#,##0"))
        End If
    Else
        If Rsh!Change_Amount > 0 Then
            Printer.Print "CHANGE";
            '--------------Normal------------------------
            Printer.CurrentX = 2000: Printer.Print "  : Rp. ";
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            CetakKanan (Format(Rsh!Change_Amount, "#,##0"))
        End If
    End If

    Printer.Print ""

    If Vsave > 0 Then
        Printer.Print "YOU SAVE";
        '--------------Normal------------------------
        Printer.CurrentX = 2000: Printer.Print "  : Rp. ";
        '--------------POSIFLEX----------------------
        'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
        CetakKanan (Format(Vsave, "#,##0"))
        Printer.Print ""
    End If

    If Trim(Rsh!card_number) <> "CM000-00000" Then
        Call MySTAR(Rsh!card_number, 0)
        'Card_No_StrukTambahan = Rsh!card_number
        Printer.Print "No MySTAR Card";
        Printer.CurrentX = 1500: Printer.Print ": " & Rsh!card_number
        Turun
        Printer.Print "Customer Name";
        Printer.CurrentX = 1500: Printer.Print ": " & Left(Star_Nm, 22)
        Turun
        Printer.Print "Issued Point";
        Printer.CurrentX = 1500: Printer.Print ": " & Get_Point(No_trans)
        Turun
        If Linked And Status <> "REPRINT" Then
            Printer.Print "Point Balance";
            Printer.CurrentX = 1500: Printer.Print ": " & Star_Pt
        End If
        Printer.Print ""
        'CetakTengah ("Tingkatkan belanja & gunakan kartu MSC")
        'CetakTengah ("Untuk memenangkan top spender")
        'Printer.Print ""
    Else
        'Card_No_StrukTambahan = "CM000-00000"
        'CetakTengah ("Not Yet A MySTAR Card member?")
        'CetakTengah ("Register Now and get more benefit")
        'Printer.Print ""
    End If

    Printer.Font.Size = 9

    CetakTengah (Tulis(11))
    CetakTengah (Tulis(12))
    CetakTengah (Tulis(13) & ", " & Tulis(14))
    CetakTengah ("NPWP/PKP No : " & Tulis(9))
    CetakTengah ("Harga sudah termasuk pajak")
    'Printer.Print ""
    'CetakTengah (Tulis(7))
    'CetakTengah (Tulis(8))
    If AcaraMKT = 1 Then
        Printer.Print ""
        CetakTengah ("Terima Kasih Anda Telah Berdonasi")
        CetakTengah ("Senilai Rp." & Format(Footer6str, "#,#00") & " Pada Program")
        CetakTengah ("*Sharing is Caring*")
        'Printer.Print ""
    End If
   
    Printer.Print ""
    CetakTengah (Tulis(3))
    CetakTengah (Tulis(4))
    CetakTengah (Tulis(5))
    Printer.Print ""

    CetakTengah ("Instagram : stardepartmentstore")
    CetakTengah ("Customer care : 0812 800 61 800 (SMS Only)")
    CetakTengah ("Website : www.stardeptstore.com")
    
    If Tulis(1) <> "" Then
        Printer.Print ""
        CetakTengah VGaris
        CetakTengah (Tulis(1))
        CetakTengah (Tulis(2))
        CetakTengah VGaris
    End If
    If Cfg_Get("Device", "Use_Barcode", App.Path & "\config.ini") = 2 Then
    Printer.Print ""
    Printer.FontName = "IDAutomationHC39M"
    'Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(No_trans)) \ 2
    Printer.Print "*" & No_trans & "*"
    End If
    Printer.Font.name = "Printer 17cpi"
    Printer.EndDoc
    Rst.Close
    Rsh.Close
End Sub

Public Sub CetakStruk_PromoEmail(Nama As String, No_trans As String, Pesan1 As String, Pesan2 As String, Pesan3 As String, Pesan4 As String, pesan5 As Integer, pesan6 As Long, Pesan As String)
Dim objPDF As New Class1
Dim FileName As String
FileName = Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & ".pdf"
    objPDF.PDFTitle = "Struk Promo"
    objPDF.PDFFileName = PathEmail & "\" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & ".pdf"
    objPDF.PDFLoadAfm = App.Path & "\Fonts"
    objPDF.PDFView = True
    objPDF.PDFBeginDoc
   
    PosY = 0
    
    objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD


    objPDF.PDFImage App.Path & "\logo.jpg", _
            8.5, 1, 4, 1, "http://www.stardeptstore.com"
    objPDF.PDFSetTextColor = vbBlack

    Call DrawBarcode(Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3), frmLogin.Picture2)
    'frmLogin.Picture2.ScaleMode = 3
    'frmLogin.Picture2.Height = frmLogin.Picture2.Height * (2.4 * 40 / frmLogin.Picture2.ScaleHeight)
    frmLogin.Picture2.FontSize = 10
    Call FileSave(frmLogin.Picture2, Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3))
    'frmLogin.Picture2.Image = Nothing
    
    If pesan5 > 0 Then
        Dim num As Integer
        Dim RsVoucher As New ADODB.Recordset
        num = 0
        If Linked Then
        Do While num < pesan5
            ConnServer.Execute "Insert Into EReceipt_Email_Detail_Voucher Select top 1 '" & No_trans & "','" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "',RTRIM(a.v_no) as v_no,'" & pesan6 & "' from newvoc a left join EReceipt_Email_Detail_Voucher b " & _
            "on a.v_no = b.KodeVoucher where V_AMT = " & pesan6 & " and " & _
            "a.v_code = '" & loc_voucher & "' and len(a.v_no) = 8 and a.V_FLAG is NULL and b.KodeVoucher is null"
            num = num + 1
        Loop
        End If
    End If
    Printer.Font.Size = 9
    objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
        objPDF.PDFImage PathEmail & "\" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & ".jpg", _
            2, 3, 18, 3, "http://www.stardeptstore.com"
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
    objPDF.PDFTextOut Pesan1, "KIRI", 1, PosYY(7)
    If Pesan2 <> "" Then
        objPDF.PDFTextOut Pesan2, "KIRI", 1, PosYY(1)
    End If
    If Pesan3 <> "" Then
        objPDF.PDFTextOut Pesan3, "KIRI", 1, PosYY(1)
    End If
    If Pesan4 <> "" Then
        objPDF.PDFTextOut Pesan4, "KIRI", 1, PosYY(1)
    End If
    If Linked Then
        ConnServer.Execute "Insert Into EReceipt_Email_Detail (Trans_Nr,Promo_Id,Namafile,Pesan,Status) values ('" & No_trans & "','" & Nama & "','" & FileName & "','" & Replace(Pesan, "'", "") & "',0)"
    End If
    
    If Linked Then
    RsVoucher.Open "Select RTRIM(KodeVoucher) KodeVoucher from EReceipt_Email_Detail_Voucher where Trans_Nr='" & No_trans & "' and promo_id = '" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
    Dim vline As Double
    Dim rpv As Variant
    Dim loop_cnt As Integer
    vline = 7
    rpv = 2
    loop_cnt = 0
    objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
    
    While Not RsVoucher.EOF
        Call DrawBarcodeV(RsVoucher!KodeVoucher, frmLogin.Picture3)
        If loop_cnt = 4 Then
            objPDF.PDFEndPage
            objPDF.PDFNewPage
            PosY = 1
            objPDF.PDFImage App.Path & "\logo.jpg", _
            8.5, 1, 4, 1, "http://www.stardeptstore.com"
            rpv = 3.7
            vline = 0
        End If
        If loop_cnt > 4 And (loop_cnt - 4) Mod 6 = 0 Then
            objPDF.PDFEndPage
            objPDF.PDFNewPage
            PosY = 1
            objPDF.PDFImage App.Path & "\logo.jpg", _
            8.5, 1, 4, 1, "http://www.stardeptstore.com"
            rpv = 3.7
            vline = 0
        End If
        Call FileSave(frmLogin.Picture3, Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "_" & RsVoucher!KodeVoucher)
        'objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
        objPDF.PDFSetTextColor = vbBlack
        objPDF.PDFTextOut "Voucher Rp." & Format(pesan6, "#,##0"), "KIRI", 3, PosYY(rpv)
        vline = vline + 4
        rpv = 5.4
        objPDF.PDFImage PathEmail & "\" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "_" & RsVoucher!KodeVoucher & ".jpg", _
            4, vline, 14, 3, "#"
            loop_cnt = loop_cnt + 1
        RsVoucher.MoveNext
    Wend
    objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
    RsVoucher.Close: Set RsVoucher = Nothing
    End If
    objPDF.PDFEndDoc
    
End Sub

Sub FileSave(Picbox As PictureBox, FileName As String)
    'This Procedure Saves the Bars to desired Format
    Dim sName, retval, retSave
    
    
    
    
    On Error GoTo ErrHandler
    
    
    
    Picbox.Picture = Picbox.Image
   
    SavePicture Picbox.Picture, PathEmail & "/" & FileName & ".bmp"
    
    PicSave.SavePicture Picbox.Picture, PathEmail & "/" & FileName & ".jpg", fmtJPEG, 70
    'Dim ImgF As WIA.ImageFile
    'Dim ImgP As WIA.ImageProcess

    'Set ImgF = New WIA.ImageFile
    'ImgF.LoadFile PathEmail & "/" & FileName & ".bmp"
    'Set ImgP = New WIA.ImageProcess
    'With ImgP
    '    .Filters.Add .FilterInfos!Convert.FilterID
    '    .Filters.Item(1).Properties!FormatID.Value = wiaFormatJPEG
    '    .Filters.Item(1).Properties!Quality.Value = 70
    '    Set ImgF = .Apply(ImgF)
    'End With
    'ImgF.SaveFile PathEmail & "/" & FileName & ".jpg"
    Exit Sub
ErrHandler:

    If Err.Number = 32755 Then ' Handle the Cancel error
        Screen.MousePointer = 0
        Exit Sub
    Else
            If Err.Number <> 0 Then MsgBox "Error saving file: " & Err.Number & " - " & Err.Description
            Screen.MousePointer = 0
    End If
    
End Sub

Public Sub CetakStrukEmail(Status As String, No_trans As String)
Dim vqty As Integer, Vsave As Long, vtotal As Long, Vbayar As Long, abc As String
Dim Rst As New ADODB.Recordset, Rsh As New ADODB.Recordset, AdaCash As Byte
Dim RsX As New ADODB.Recordset
Dim pageCnt As Integer

    Dim objPDF As New Class1
    objPDF.PDFTitle = "Struk"
    objPDF.PDFFileName = PathEmail & "\" & Trim(No_trans) & ".pdf"
    objPDF.PDFLoadAfm = App.Path & "\Fonts"
    objPDF.PDFView = True
    objPDF.PDFBeginDoc
    'cek cash terakhir
    
    'objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
    'objPDF.PDFFormatPage = FORMAT_A4
    'objPDF.PDFOrientation = ORIENT_PORTRAIT
    'objPDF.PDFSetUnit = UNIT_PT
    
    PosY = 0
    pageCnt = 1
    
    objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
  
    
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
    
    objPDF.PDFImage App.Path & "\logo.jpg", _
            8.5, 1, 4, 1, "http://www.stardeptstore.com"
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFTextOut "DEPARTMENT STORE", "TENGAH", 0, PosYY(4)

    Printer.Font.Size = 9
    objPDF.PDFTextOut (Tulis(10)), "TENGAH", 0, PosYY(1)
    
    objPDF.PDFTextOut "No. " & No_trans, "KIRI", 1, PosYY(2)
    objPDF.PDFTextOut VShift & "-" & VKasir_ID & "/" & Left(Trim(VKasir_Nm), 14) & "     " & Format(Rsh!Transaction_Date, "dd/mm/yyyy") & " " & Rsh!Transaction_Time, "KIRI", 1, PosYY(1)


    If Status <> "SALES" Then

        Select Case Status
        Case "REFUND"
            objPDF.PDFTextOut "REFUND TRANSACTION", "TENGAH", 0, PosYY(1)
        Case "REPRINT"
            objPDF.PDFTextOut "R E P R I N T", "TENGAH", 0, PosYY(1)
        End Select
    End If

    Rst.Open "select Seq, sd.PLU, Item_Description, Price, Qty, Discount_Percentage, Discount_Amount, " & _
            "ExtraDisc_Pct, ExtraDisc_Amt, net_price, brand from sales_transaction_details sd inner join item_master im " & _
            "on sd.plu=im.plu where transaction_number='" & No_trans & "' order by seq", ConnLocal, adOpenStatic, adLockReadOnly
    objPDF.PDFTextOut "", "KIRI", 0, PosYY(1)
    While Not Rst.EOF

        objPDF.PDFTextOut Left(Trim(Rst!plu) & " " & Trim(Rst!Item_Description), 42), "KIRI", 1, PosYY(1)
        abc = "  " & Rst!Qty & "x" & Format(Rst!price, "#,##0") & " " & IIf(Rst!Brand = "No Brand", " ", Rst!Brand)
        abc = Left(abc, 30)

        objPDF.PDFTextOut abc, "KIRI", 1, PosYY(1)
        '---------------------Normal----------------
        objPDF.PDFTextOut Format(Rst!Net_Price, "#,##0"), "KANAN", 0, PosYY(0)
        '---------------------POSIFLEX----------------
        'CetakKonon (Format(Rst!Net_Price, "#,##0"))
        If Rst!Discount_Percentage <> 0 Then
            objPDF.PDFTextOut "  Disc. " & Rst!Discount_Percentage & "% = " & Format(Rst!Discount_Amount, "#,##0"), "KIRI", 1, PosYY(1)
        End If

        If Rst!ExtraDisc_pct <> 0 Then
            Turun
            objPDF.PDFTextOut "  Extra " & Rst!ExtraDisc_pct & "% = " & Format(Rst!ExtraDisc_amt, "#,##0"), "KIRI", 1, PosYY(1)
        End If

        vqty = vqty + Rst!Qty
        vtotal = vtotal + Rst!Net_Price
        Vsave = Vsave + Rst!Discount_Amount + Rst!ExtraDisc_amt
        Rst.MoveNext
        If PosY > 32 Then
            objPDF.PDFEndPage
            objPDF.PDFNewPage
            PosY = 1
    End If
    Wend

    objPDF.PDFTextOut "Total   " & Right("   " & vqty, 4) & " item(s)   ", "KIRI", 1, PosYY(1)
    '--------------Normal------------------------
    objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(0)
    '--------------POSIFLEX----------------------
    'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
    objPDF.PDFTextOut Kanan(14, vtotal), "KANAN", 0, PosYY(0)


    While Not Rsh.EOF
        If Trim(Rsh!Credit_Card_No) <> "" Then
            If Len(Trim(Rsh!Credit_Card_No)) = 16 Then
                objPDF.PDFTextOut Left(Rsh!Credit_Card_No, 7) & "XXXXXXXXX", "KIRI", 1, PosYY(1)
            Else
                objPDF.PDFTextOut Left(Rsh!Credit_Card_No, 20), "KIRI", 1, PosYY(1)
            End If
            If Trim(Rsh!credit_card_name) <> "" Then objPDF.PDFTextOut Left(Rsh!credit_card_name, 40), "KIRI", 1, PosYY(1)
        End If

        abc = Left(Rsh!Description & Space(22), 22)
        If Rsh!Payment_Types > 30 Then abc = Left(Rsh!credit_card_name & Space(24), 24)

        If Trim(Rsh!Description) = "CASH" Then
            objPDF.PDFTextOut "CASH", "KIRI", 1, PosYY(1)
            '--------------Normal------------------------
            objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(0)
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            If Rsh!urut = AdaCash Then
                objPDF.PDFTextOut Format((Rsh!paid_amount + Rsh!Change_Amount), "#,##0"), "KANAN", 0, PosYY(0)
            Else
                objPDF.PDFTextOut Format(Rsh!paid_amount, "#,##0"), "KANAN", 0, PosYY(0)
            End If
        Else
            objPDF.PDFTextOut abc, "KIRI", 1, PosYY(1)
            '--------------Normal------------------------
            objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(0)
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            objPDF.PDFTextOut Format(Rsh!paid_amount, "#,##0"), "KANAN", 0, PosYY(0)
        End If
        Rsh.MoveNext
    Wend

    If PosY > 32 Then
        objPDF.PDFEndPage
        objPDF.PDFNewPage
        PosY = 1
    End If
        
    Rsh.MoveFirst

    If Status = "REFUND" Then
        If Rsh!Change_Amount < 0 Then
            objPDF.PDFTextOut "CHANGE", "KIRI", 1, PosYY(1)
            objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(1)
            objPDF.PDFTextOut Format(Rsh!Change_Amount, "#,##0"), "KANAN", 0, PosYY(0)
        End If
    Else
        If Rsh!Change_Amount > 0 Then
            objPDF.PDFTextOut "CHANGE", "KIRI", 1, PosYY(1)
            '--------------Normal------------------------
            objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(1)
            '--------------POSIFLEX----------------------
            'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
            objPDF.PDFTextOut Format(Rsh!Change_Amount, "#,##0"), "KANAN", 0, PosYY(0)
        End If
    End If

    If Vsave > 0 Then
        objPDF.PDFTextOut "YOU SAVE", "KIRI", 1, PosYY(1)
        '--------------Normal------------------------
        objPDF.PDFTextOut "  : Rp. ", "KIRI", 7, PosYY(0)
        '--------------POSIFLEX----------------------
        'Printer.CurrentX = 1800: Printer.Print "  : Rp. ";
        objPDF.PDFTextOut Format(Vsave, "#,##0"), "KANAN", 0, PosYY(0)

    End If

    If Trim(Rsh!card_number) <> "CM000-00000" Then
        Call MySTAR(Rsh!card_number, 0)
        objPDF.PDFTextOut "No MySTAR Card", "KIRI", 1, PosYY(2)
        objPDF.PDFTextOut ": " & Rsh!card_number, "KIRI", 6, PosYY(0)
        objPDF.PDFTextOut "Customer Name", "KIRI", 1, PosYY(1)
        objPDF.PDFTextOut ": " & Left(Star_Nm, 22), "KIRI", 6, PosYY(0)
        objPDF.PDFTextOut "Issued Point", "KIRI", 1, PosYY(1)
        objPDF.PDFTextOut ": " & Get_Point(No_trans), "KIRI", 6, PosYY(0)
        If Linked And Status <> "REPRINT" Then
            objPDF.PDFTextOut "Point Balance", "KIRI", 1, PosYY(1)
            objPDF.PDFTextOut ": " & Star_Pt, "KIRI", 6, PosYY(0)
        End If
    Else

    End If
    
    If PosY > 32 Then
        objPDF.PDFEndPage
        objPDF.PDFNewPage
        PosY = 1
    End If
    
    objPDF.PDFTextOut Tulis(11), "TENGAH", 0, PosYY(2)
    objPDF.PDFTextOut Tulis(12), "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut Tulis(13) & ", " & Tulis(14), "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut "NPWP/PKP No : " & Tulis(9), "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut "Harga sudah termasuk pajak", "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut Tulis(7), "TENGAH", 0, PosYY(2)
    objPDF.PDFTextOut Tulis(8), "TENGAH", 0, PosYY(1)
    If AcaraMKT = 1 Then
        objPDF.PDFTextOut "Terima Kasih Telah Berpartisipasi", "TENGAH", 0, PosYY(2)
        objPDF.PDFTextOut "Dalam Program Charity 'Pink Power'", "TENGAH", 0, PosYY(1)
        objPDF.PDFTextOut "Donasi Anda Sebesar Rp." & Format(Footer6str, "#,#00"), "TENGAH", 0, PosYY(1)
    End If
   
    If PosY > 32 Then
        objPDF.PDFEndPage
        objPDF.PDFNewPage
        PosY = 1
    End If
   
    objPDF.PDFTextOut Tulis(3), "TENGAH", 0, PosYY(2)
    objPDF.PDFTextOut Tulis(4), "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut Tulis(5), "TENGAH", 0, PosYY(1)


    objPDF.PDFTextOut "Facebook / Twitter stardeptstore", "TENGAH", 0, PosYY(2)
    objPDF.PDFTextOut "Customer care : 0812 800 61 800 (SMS Only)", "TENGAH", 0, PosYY(1)
    objPDF.PDFTextOut "Website : http://www.stardeptstore.com", "TENGAH", 0, PosYY(1)
    
    If Tulis(1) <> "" Then
        Printer.Print ""
        objPDF.PDFTextOut VGaris, "TENGAH", 0, PosYY(2)
        objPDF.PDFTextOut Tulis(1), "TENGAH", 0, PosYY(1)
        objPDF.PDFTextOut Tulis(2), "TENGAH", 0, PosYY(1)
        objPDF.PDFTextOut VGaris, "TENGAH", 0, PosYY(1)
    End If
    Rst.Close
    objPDF.PDFEndDoc
    Rsh.Close
End Sub

Public Sub Kirim_Promo_Mobile(Nama As String, Tipe As Integer, No_Kartu As String, No_trans As String, Pesan1 As String, Pesan2 As String, Pesan3 As String, Pesan4 As String, Count As Integer, Msg As String)
If No_Kartu = "CM000-00000" Then
    Call CetakStruk_Promo(Nama, No_trans, 0, 0, Msg)
    Exit Sub
End If
Dim RsX As New ADODB.Recordset
RsX.Open "SELECT Nominal From mobile_tipe_promo where Promo_Id = '" & Nama & "' And aktif = 1", ConnServer, adOpenStatic, adLockReadOnly
Dim I As Integer
If Not RsX.EOF Then
ConnServer.Execute "Insert Into Promo_Mobile_Hdr (Promo_Id,Transaction_Number,Card_Number,Tipe,Msg1,Msg2,Msg3,Msg4,Status,Create_Date,Last_Update)" & _
   " Values ('" & Nama & "','" & No_trans & "','" & No_Kartu & "','" & Tipe & "','" & Replace(Pesan1, "'", "") & "','" & Replace(Pesan2, "'", "") & "', " & _
   "'" & Replace(Pesan3, "'", "") & "','" & Replace(Pesan4, "'", "") & "',1,getdate(),getdate())"
For I = 1 To Count
   ConnServer.Execute "Insert Into Promo_Mobile_Dtl (Transaction_Number,Seq,Nominal)" & _
   " Values ('" & No_trans & "'," & I & "," & RsX!Nominal & ")"
Next
End If

RsX.Close: Set RsX = Nothing

End Sub

Public Sub CetakStrukPayPoint(Card_Nr As String, Card_Name As String, No_trans As String)
Dim RsX As New ADODB.Recordset
Dim Rst As New ADODB.Recordset
Dim AdaPayPoint As Integer
Dim SisaPayPoint As Integer
    'cek pembayaran point
    RsX.Open "SELECT Sum(Paid_Amount) As Amount From Paid " & _
             "WHERE (Transaction_Number = '" & No_trans & "') AND (Payment_Types = '5') " & _
             "", ConnServer, adOpenStatic, adLockReadOnly
    
    If Not RsX.EOF Then
        If IsNull(RsX!Amount) Then
            AdaPayPoint = 0
        Else
            AdaPayPoint = RsX!Amount / VHargaPoint
        End If
    Else
        AdaPayPoint = 0
    End If
    
    RsX.Close: Set RsX = Nothing
    
    If AdaPayPoint = 0 Then
    Exit Sub
    End If
    
    Rst.Open "select card_point from card where card_nr = '" & Card_Nr & "'", ConnServer, adOpenStatic, adLockReadOnly
    If Not Rst.EOF Then
        SisaPayPoint = Rst!card_point
    Else
        SisaPayPoint = 0
    End If
    If Cfg_Get("Device", "Use_Logo", App.Path & "\config.ini") = 0 Then
    CetakTengah ("STAR DEPARTMENT STORE")
    Else
    CetakTengah ("DEPARTMENT STORE")
    End If
    
    
    Turun
    Printer.Font.Size = 9
    CetakTengah (Tulis(10))
    
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Turun
    Printer.Print ""
    Printer.Print "POINT REWARD REDEMPTION"
    Turun
    Printer.Print "No.";
    Printer.CurrentX = 1500: Printer.Print ": " & No_trans
    Turun
    Printer.Print "No MySTAR Card";
    Printer.CurrentX = 1500: Printer.Print ": " & Card_Nr
    Turun
    Printer.Print "Customer Name";
    Printer.CurrentX = 1500: Printer.Print ": " & Left(Star_Nm, 22)
    Turun
    Printer.Print "Claim Date";
    Printer.CurrentX = 1500: Printer.Print ": " & Now()
    Turun
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.Print "Claim Point";
    Printer.CurrentX = 1500: Printer.Print ": " & AdaPayPoint
    Turun
    Printer.Print "Point Balance";
    Printer.CurrentX = 1500: Printer.Print ": " & Star_Pt
    Printer.Print ""
    Printer.EndDoc
    Rst.Close
End Sub

Public Sub CetakStrukEVoucherOnline(Nama As String, No_trans As String, pesan5 As Integer, pesan6 As Long)
Dim RsVoucher As New ADODB.Recordset
    If pesan5 > 0 Then
        Dim num As Integer
        num = 0
        If Linked Then
        Do While num < pesan5
            ConnServer.Execute "Insert Into EReceipt_Email_Detail_Voucher Select top 1 '" & No_trans & "','" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "',RTRIM(a.v_no) as v_no,'" & pesan6 & "' from newvoc a left join EReceipt_Email_Detail_Voucher b " & _
            "on a.v_no = b.KodeVoucher where V_AMT = " & pesan6 & " and " & _
            "a.v_code = '" & loc_voucher & "' and len(a.v_no) = 8 and a.V_FLAG is NULL and b.KodeVoucher is null"
            num = num + 1
        Loop
        End If
    End If

    If Linked Then
        ConnServer.Execute "Insert Into EReceipt_Email_Detail (Trans_Nr,Promo_Id,Namafile,Pesan,Status) values ('" & No_trans & "','" & Nama & "','','',0)"
    End If
    
    If Linked Then
    RsVoucher.Open "Select RTRIM(KodeVoucher) KodeVoucher,Nominal from EReceipt_Email_Detail_Voucher where Trans_Nr='" & No_trans & "' and promo_id = '" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
    End If
    If Cfg_Get("Device", "Use_Logo", App.Path & "\config.ini") = 0 Then
    CetakTengah ("STAR DEPARTMENT STORE")
    Else
    CetakTengah ("DEPARTMENT STORE")
    End If
    
    Turun
    Printer.Font.Size = 9
    CetakTengah (Tulis(10))
    
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Turun
    Printer.Print ""
    Printer.Print "E-VOUCHER"
    Turun
    Printer.Print "Print Date";
    Printer.CurrentX = 1500: Printer.Print ": " & Now()
    Turun
    Printer.Print "Trans No";
    Printer.CurrentX = 1500: Printer.Print ": " & No_trans
    Turun
    
    While Not RsVoucher.EOF
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.FontName = "IDAutomationHC39M"
    'Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(No_trans)) \ 2
    Printer.FontSize = 11
    Printer.Print "*" & Trim(RsVoucher!KodeVoucher) & "*"
    Turun
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "Voucher Code";
    Printer.CurrentX = 1500: Printer.Print ": " & Trim(RsVoucher!KodeVoucher)
    Turun
    Printer.Print "Voucher Amount";
    Printer.CurrentX = 1500: Printer.Print ": " & Format(RsVoucher!Nominal, "#,##0")
    Turun
    RsVoucher.MoveNext
    Wend
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.Print ""
    Printer.EndDoc
    RsVoucher.Close: Set RsVoucher = Nothing
End Sub

Public Sub CetakStrukEVoucher(No_trans As String)
Dim RsX As New ADODB.Recordset
Dim TotalV As Long


    'cek pembayaran Voucher Poll
    RsX.Open "SELECT Credit_Card_No, Paid_Amount As Amount From Paid " & _
             "WHERE (Transaction_Number = '" & No_trans & "') AND (Payment_Types = '8') AND (LEN(Credit_Card_No) = 8)" & _
             "", ConnServer, adOpenStatic, adLockReadOnly
    
    If RsX.EOF Then
        RsX.Close: Set RsX = Nothing
        Exit Sub
    End If
    TotalV = 0
    If Cfg_Get("Device", "Use_Logo", App.Path & "\config.ini") = 0 Then
    CetakTengah ("STAR DEPARTMENT STORE")
    Else
    CetakTengah ("DEPARTMENT STORE")
    End If
    
     Turun
    Printer.Font.Size = 9
    CetakTengah (Tulis(10))
    
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Turun
    Printer.Print ""
    Printer.Print "E-VOUCHER REDEMPTION"
    Turun
    Printer.Print "Claim Date";
    Printer.CurrentX = 1500: Printer.Print ": " & Now()
    Turun
    Printer.Print "Trans No";
    Printer.CurrentX = 1500: Printer.Print ": " & No_trans
    Turun
    
    While Not RsX.EOF
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.FontName = "IDAutomationHC39M"
    'Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(No_trans)) \ 2
    Printer.FontSize = 11
    Printer.Print "*" & Trim(RsX!Credit_Card_No) & "*"
    Turun
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "Voucher Code";
    Printer.CurrentX = 1500: Printer.Print ": " & Trim(RsX!Credit_Card_No)
    Turun
    Printer.Print "Voucher Amount";
    Printer.CurrentX = 1500: Printer.Print ": " & Format(RsX!Amount, "#,##0")
    TotalV = TotalV + RsX!Amount
    Turun
    RsX.MoveNext
    Wend
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.Print "Total Amount";
    Printer.CurrentX = 1500: Printer.Print ": " & Format(TotalV, "#,##0")
    Turun
    Printer.Print ""
    Printer.EndDoc
    RsX.Close: Set RsX = Nothing
End Sub

Public Sub CetakStrukDANA(No_trans As String)
On Error GoTo ErrH
Dim RsX As New ADODB.Recordset
Dim TotalV As Long


    'cek pembayaran DANA
    RsX.Open "SELECT Credit_Card_No, Paid_Amount As Amount From Paid " & _
             "WHERE (Transaction_Number = '" & No_trans & "') AND (Payment_Types = '22') " & _
             "", ConnLocal, adOpenStatic, adLockReadOnly
    
    If RsX.EOF Then
        RsX.Close: Set RsX = Nothing
        Exit Sub
    End If
    TotalV = 0
    If Cfg_Get("Device", "Use_Logo", App.Path & "\config.ini") = 0 Then
    CetakTengah ("STAR DEPARTMENT STORE")
    Else
    CetakTengah ("DEPARTMENT STORE")
    End If
    
     Turun
    Printer.Font.Size = 9
    CetakTengah (Tulis(10))
    
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Turun
    Printer.Print ""
    Printer.Print "DANA PAYMENT"
    Turun
    Printer.Print "Payment Date";
    Printer.CurrentX = 1500: Printer.Print ": " & Now()
    Turun
    Printer.Print "Trans No";
    Printer.CurrentX = 1500: Printer.Print ": " & No_trans
    Turun
    
    While Not RsX.EOF
    Printer.Print "-----------------------------------------------------"
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "DANA Code";
    Printer.CurrentX = 1500: Printer.Print ": " & Trim(RsX!Credit_Card_No)
    Turun
    Printer.Print "DANA Amount";
    Printer.CurrentX = 1500: Printer.Print ": " & Format(RsX!Amount, "#,##0")
    TotalV = TotalV + RsX!Amount
    Turun
    RsX.MoveNext
    Wend
    Printer.Print "-----------------------------------------------------"
    Turun
    Printer.Print "Total Amount";
    Printer.CurrentX = 1500: Printer.Print ": " & Format(TotalV, "#,##0")
    Turun
    Printer.Print ""
    Printer.EndDoc
    RsX.Close: Set RsX = Nothing
    Exit Sub
ErrH:
    RsX.Close: Set RsX = Nothing
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog("Cetak DANA " & Err.Description & " " & Err.Number)
End Sub


Public Sub CetakStruk_Promo(Nama As String, No_trans As String, pesan5 As Integer, pesan6 As Long, Pesan As String)
    Dim RsVoucher As New ADODB.Recordset
    Dim ikutvoucher As Integer
    If Linked Then
    RsVoucher.Open "Select ISNULL(Gift_Voucher,0) Gift_Voucher from Promo_hdr where promo_id = '" & Nama & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
    ikutvoucher = 0
    If Not RsVoucher.EOF Then
        ikutvoucher = RsVoucher!Gift_Voucher
    End If
    If ikutvoucher = 1 Then
        ConnServer.Execute "Insert Into EReceipt_Email (Email,Trans_Nr,Gender,Nama,Tanggal,Reprint,status) values ('','" & No_trans & "','" & Star_Gender & "', '',getdate(),0, 0)"
        Call CetakStrukEVoucherOnline(Nama, No_trans, pesan5, pesan6)
        GoTo 1
    Else
        RsVoucher.Close: Set RsVoucher = Nothing
    End If
    End If
    If Cfg_Get("Device", "Use_Barcode", App.Path & "\config.ini") = 1 Then
    Printer.Print ""
    Printer.FontName = "IDAutomationHC39M"
    'Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(No_trans)) \ 2
    Printer.FontSize = 11
    Printer.Print "*" & Left(Nama, 3) & Mid(No_trans, 4, 4) & Mid(No_trans, 9, 4) & Mid(No_trans, 15, 2) & Mid(No_trans, 19, 3) & "*"
    Printer.FontSize = 8
    Printer.Font.name = "Printer 17cpi"
    Printer.Print ""
    Else
    Printer.Print VGaris
    Printer.Print "TRANS#";
    Printer.CurrentX = 800: Printer.Print ": "; No_trans
    Printer.Print "CASHIER";
    Printer.CurrentX = 800: Printer.Print ": "; VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Printer.Print "TIME";
    Printer.CurrentX = 800: Printer.Print ": "; Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Printer.Print VGaris
    End If
    'Printer.Print "No MySTAR Card";
    'Printer.CurrentX = 1500: Printer.Print ": " & Card_No_StrukTambahan
    Printer.Print Pesan
    Printer.Print VGaris
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.EndDoc
1:
End Sub

Public Sub CetakPesan(Status As String, No_trans As String)
    'Printer.Print Tulis(11) 'nama pt
    Printer.Print Tulis(10)
    Printer.Print "TRANS#";
    Printer.CurrentX = 800: Printer.Print ": "; No_trans
    Printer.Print "CASHIER";
    Printer.CurrentX = 800: Printer.Print ": "; VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Printer.Print "TIME";
    Printer.CurrentX = 800: Printer.Print ": "; Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Select Case Status
    Case "HOLD"
        Printer.Print "HOLD TRANSACTION"
    Case "CANCEL"
        Printer.Print "CANCEL TRANSACTION"
    End Select
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.EndDoc
End Sub

Public Sub CetakBegin()
    'Printer.PrintQuality = vbPRPQDraft
    'Printer.PaintPicture LoadPicture("star.jpg"), 0, 0, 1104.8, 533.2
    'Printer.CurrentY = Printer.CurrentY + 600
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "POS BEGIN... "
    Printer.Print "NPWP";
    Printer.CurrentX = 1000: Printer.Print ": " & Tulis(9)
    Printer.Print "REGISTER";
    Printer.CurrentX = 1000: Printer.Print ": "; VReg_ID
    Printer.Print "CASHIER";
    Printer.CurrentX = 1000: Printer.Print ": " & VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Printer.Print "TIME";
    Printer.CurrentX = 1000: Printer.Print ": "; Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Printer.EndDoc
End Sub

Public Sub CetakValid(No_trans As String, brs1 As String, brs2 As String)
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 8
    Printer.Print "TRANS#";
    Printer.CurrentX = 800: Printer.Print ": "; No_trans
    Printer.Print "CASHIER";
    Printer.CurrentX = 800: Printer.Print ": "; VShift & " - " & VKasir_ID & "/" & VKasir_Nm
    Printer.Print "TIME";
    Printer.CurrentX = 800: Printer.Print ": "; Format(Now(), "dd/mmm/yyyy HH:MM:SS")
    Printer.Print brs1
    Printer.Print brs2
    'Printer.Print Left(brs2, 40)
    'Printer.Print Mid(brs2, 21, Len(brs2) - 40)
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.Print "": Printer.Print "": Printer.Print ""
    Printer.EndDoc
End Sub

Public Sub OpenLaci(Tipe As Byte)
    Call Open_OPOS_Drawer
    If Tipe = 1 Then
'        Printer.Print Chr(27) + Chr(64); 'reset printer
'        Printer.Print Chr(27) & Chr(33) & Chr(1); '10cpi
        Printer.Print "CASHIER";
        Printer.CurrentX = 1000: Printer.Print ": "; VShift & " - " & VKasir_ID & "/" & VKasir_Nm
        Printer.Print "TIME";
        Printer.CurrentX = 1000: Printer.Print ": "; Format(Now(), "dd/mmm/yyyy HH:MM:SS")
        Printer.Print "OPEN DRAWER"
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.EndDoc
    End If
End Sub
    
Private Sub Open_OPOS_Drawer()
'    frmMain.OPOSCashDrawer1.Open ("LACI")
'    frmMain.OPOSCashDrawer1.ClaimDevice (0)
'    frmMain.OPOSCashDrawer1.PowerNotify = 1
'    frmMain.OPOSCashDrawer1.DeviceEnabled = True
'    frmMain.OPOSCashDrawer1.OpenDrawer

If Not frmMain.OPOSCashDrawer1.OpenDrawer Then
frmMain.OPOSCashDrawer1.Open ("LACI")
frmMain.OPOSCashDrawer1.ClaimDevice (1000)
frmMain.OPOSCashDrawer1.DeviceEnabled = True
frmMain.OPOSCashDrawer1.OpenDrawer
frmMain.OPOSCashDrawer1.DeviceEnabled = False
frmMain.OPOSCashDrawer1.ReleaseDevice
frmMain.OPOSCashDrawer1.Close
End If

End Sub

'Private Sub Open_OPOS_Drawer()
'Open "lpt1" For Output As #1
'Print #1, Chr$(&H1B); "p"; Chr$(0); Chr$(100); Chr$(250);
'Print #1, Chr$(&H1B); "u"; Chr$(0);
'Close #1
'End Sub



Public Sub XRead()
Dim RsBayar As New ADODB.Recordset, Rs As New ADODB.Recordset, RsPlastic As New ADODB.Recordset
Dim Jual As Long, diskon As Long, Retur As Long, Batal As Long, Modal As Long, Jumlah As Long
    
    Call OpenLaci(0)
    If Cfg_Get("Device", "X_ReadPrint", App.Path & "\config.ini") = 1 Then
        Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Nilai, isnull(SUM(Total_discount),0) AS Potong " & _
                "FROM Sales_Transactions WHERE Status = '00' and substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
                "' and Transaction_Number in (select transaction_number from paid where Paid.Shift = '" & VShift & "') ", ConnLocal, adOpenForwardOnly, adLockReadOnly
        Jual = Rs!Nilai
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
        Batal = Rs!Nilai
        Rs.Close
    
        Rs.Open "SELECT Modal From Cash WHERE (Branch_ID = '" & VBranch_ID & "') AND (Datetime = '" & _
            Format(GetSrvDate, "YYYY-MM-DD") & "') AND (Shift = " & VShift & ")", ConnLocal, adOpenForwardOnly, adLockReadOnly
        Modal = Rs!Modal
        Rs.Close: Set Rs = Nothing
    
        RsBayar.Open "SELECT Payment_Types.Description, SUM(Paid.Paid_Amount) AS Nilai " & _
            "FROM Paid INNER JOIN Payment_Types ON Paid.Payment_Types = Payment_Types.Payment_Types " & _
            "WHERE (Paid.Shift = '" & VShift & "') and substring(transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
            "' GROUP BY Payment_Types.seq, Payment_Types.Description order by Payment_Types.seq", ConnLocal, adOpenForwardOnly, adLockReadOnly
        Printer.Font.Size = 8
        Printer.Print "X-Reading Shift";
        Printer.CurrentX = 1500: Printer.Print ": " & VShift & " " & frmMain.lblline
        Printer.Print "BRANCH";
        Printer.CurrentX = 1500: Printer.Print ": " & Tulis(10)
        Printer.Print "REGISTER";
        Printer.CurrentX = 1500: Printer.Print ": " & VReg_ID
        Printer.Print "CASHIER";
        Printer.CurrentX = 1500: Printer.Print ": " & VKasir_ID & "/" & VKasir_Nm
        Printer.Print "TIME";
        Printer.CurrentX = 1500: Printer.Print ": " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
        Printer.Print VGaris
        Printer.Print "MODAL";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Modal, "#,##0") & "    "
    
        While Not RsBayar.EOF
            Printer.Print Left(RsBayar!Description & Space(20), 20);
            Printer.CurrentX = 1800: Printer.Print ": Rp. ";
            CetakKanan Format(RsBayar!Nilai, "#,##0") & "    "
    '        Printer.Print Left(RsBayar!Description & Space(20), 20) & "   : Rp. " & Kanan(14, RsBayar!nilai)
            Jumlah = Jumlah + RsBayar!Nilai
            RsBayar.MoveNext
        Wend
    
        Printer.Print VGaris
        Printer.Print "TOTAL";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jumlah, "#,##0") & "    "
        Printer.Print "OVER VOUCHER";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jumlah - Jual, "#,##0") & "    "
        Printer.Print VGaris
        Printer.Print "X Reading";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jual, "#,##0") & "    "
        Printer.Print ""
        Printer.Print "DISCOUNT";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(diskon, "#,##0") & "    "
        Printer.Print "RETURN";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Retur, "#,##0") & "    "
        Printer.Print "VOID";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Batal, "#,##0") & "    "
        Printer.Print VGaris
        'plastic
        RsPlastic.Open "SELECT Sales_Transaction_Details.Item_Description, Count(Sales_Transaction_Details.Item_Description) AS Nilai " & _
            "FROM Sales_Transactions INNER JOIN Sales_Transaction_Details ON Sales_Transactions.Transaction_Number = Sales_Transaction_Details.Transaction_Number " & _
            " INNER JOIN Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU WHERE (Sales_Transactions.Status = '00') and substring(Sales_Transactions.transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
            "' AND BURUI = 'NMD31ZZZ9' GROUP BY Sales_Transaction_Details.Item_Description order by Sales_Transaction_Details.Item_Description", ConnLocal, adOpenForwardOnly, adLockReadOnly
        
        While Not RsPlastic.EOF
            Printer.Print Left(RsPlastic!Item_Description & Space(20), 20);
            Printer.CurrentX = 1800: Printer.Print ":";
            CetakKanan Format(RsPlastic!Nilai, "#,##0") & "    "
            RsPlastic.MoveNext
        Wend
        
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.EndDoc
        'update table cash
        RsBayar.Close: Set RsBayar = Nothing
        RsPlastic.Close: Set RsPlastic = Nothing
    End If
    
    
    ConnLocal.Execute "update cash_register set shift='2' WHERE Branch_ID = '" & VBranch_ID & _
                      "' AND cash_register_id='" & VReg_ID & "'"
    If Linked Then ConnServer.Execute "update cash_register set shift='2' WHERE Branch_ID = '" & VBranch_ID & _
                      "' AND cash_register_id='" & VReg_ID & "'"

End Sub

Public Sub ZReset()
Dim RsBayar As New ADODB.Recordset, Rs As New ADODB.Recordset, x As Byte, RsPlastic As New ADODB.Recordset
Dim Jual As Long, diskon As Long, Retur As Long, Batal As Long, Modal As Long, Jumlah As Long
    
    Call OpenLaci(0)
    If Cfg_Get("Device", "X_ReadPrint", App.Path & "\config.ini") = 1 Then
        Rs.Open "SELECT isnull(SUM(Net_amount),0) AS Nilai, isnull(SUM(Total_discount),0) AS Potong " & _
                "FROM Sales_Transactions WHERE Status = '00' and substring(transaction_number, 9,8)='" & _
                Format(GetSrvDate, "DDMMYYYY") & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
        Jual = Rs!Nilai
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
        Batal = Rs!Nilai
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
        Printer.Font.Size = 8
        Printer.Print "Z-Reset Shift";
        Printer.CurrentX = 1500: Printer.Print ": " & VShift & " " & frmMain.lblline
        Printer.Print "BRANCH";
        Printer.CurrentX = 1500: Printer.Print ": " & Tulis(10)
        Printer.Print "REGISTER";
        Printer.CurrentX = 1500: Printer.Print ": " & VReg_ID
        Printer.Print "CASHIER";
        Printer.CurrentX = 1500: Printer.Print ": " & VKasir_ID & "/" & VKasir_Nm
        Printer.Print "TIME";
        Printer.CurrentX = 1500: Printer.Print ": " & Format(Now(), "dd/mmm/yyyy HH:MM:SS")
        Printer.Print VGaris
        Printer.Print "MODAL";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Modal, "#,##0") & "    "
    
        If Not (RsBayar.EOF And RsBayar.BOF) Then RsBayar.MoveFirst
        While Not RsBayar.EOF
            Printer.Print Left(RsBayar!Description & Space(20), 20);
            Printer.CurrentX = 1800: Printer.Print ": Rp. ";
            CetakKanan Format(RsBayar!Nilai, "#,##0") & "    "
    '        Printer.Print Left(RsBayar!Description & Space(20), 20) & "   : Rp. " & Kanan(14, RsBayar!nilai)
            Jumlah = Jumlah + RsBayar!Nilai
            RsBayar.MoveNext
        Wend
    
        Printer.Print VGaris
        Printer.Print "TOTAL";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jumlah, "#,##0") & "    "
        Printer.Print "OVER VOUCHER";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jumlah - Jual, "#,##0") & "    "
        Printer.Print VGaris
        Printer.Print "Z Reset";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Jual, "#,##0") & "    "
        Printer.Print ""
        Printer.Print "DISCOUNT";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(diskon, "#,##0") & "    "
        Printer.Print "RETURN";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Retur, "#,##0") & "    "
        Printer.Print "VOID";
        Printer.CurrentX = 1800: Printer.Print ": Rp. ";
        CetakKanan Format(Batal, "#,##0") & "    "
    Printer.Print VGaris
    
        'plastic
        RsPlastic.Open "SELECT Sales_Transaction_Details.Item_Description, Count(Sales_Transaction_Details.Item_Description) AS Nilai " & _
            "FROM Sales_Transactions INNER JOIN Sales_Transaction_Details ON Sales_Transactions.Transaction_Number = Sales_Transaction_Details.Transaction_Number " & _
            " INNER JOIN Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU WHERE (Sales_Transactions.Status = '00') and substring(Sales_Transactions.transaction_number, 9,8)='" & Format(GetSrvDate, "DDMMYYYY") & _
            "' AND BURUI = 'NMD31ZZZ9' GROUP BY Sales_Transaction_Details.Item_Description order by Sales_Transaction_Details.Item_Description", ConnLocal, adOpenForwardOnly, adLockReadOnly
        
        While Not RsPlastic.EOF
            Printer.Print Left(RsPlastic!Item_Description & Space(20), 20);
            Printer.CurrentX = 1800: Printer.Print ":";
            CetakKanan Format(RsPlastic!Nilai, "#,##0") & "    "
            RsPlastic.MoveNext
        Wend
        RsPlastic.Close: Set RsPlastic = Nothing
        
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.Print "": Printer.Print "": Printer.Print ""
        Printer.EndDoc
    
        Next x
        RsBayar.Close: Set RsBayar = Nothing
        
    End If
    Dim Fso As New FileSystemObject
    Dim fil As File
    For Each fil In Fso.GetFolder(PathEmail & "\BACKUP").Files
        Debug.Print
        Dim TabFile2() As String
        TabFile2 = Split(fil.name, ".")
        Kill PathEmail & "\BACKUP" & "\" & fil.name
    Next
    'Call SQLQuery("update cash_register set shift='1', last_reset_date=reset_date, reset_date=getdate(), zreset_status=1 WHERE Branch_ID = '" & _
                VBranch_ID & "' AND cash_register_id='" & VReg_ID & "'")
    
    'ConnLocal.Execute "update branches set date_yesterday=date_current, date_current=getdate() " & _
                     "WHERE Branch_ID = '" & VBranch_ID & "'"

    ConnLocal.Execute "exec spp_ZresetLocal '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "'"
    ConnLocal.Execute "exec spp_ZresetServer '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "',''"
    ConnLocal.Execute "exec spp_DeleteTrans"
    
    If Linked Then ConnServer.Execute "exec spp_ZresetServer '" & VBranch_ID & "', '" & VReg_ID & "', '" & Format(GetSrvDate, "YYYY-MM-DD") & "',''"
    

End Sub

Public Sub CetakData()
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "-------------------------------------------------------------------"
    Printer.Print "Card Number";
    Printer.CurrentX = 1500: Printer.Print ": " & Star_No:
    Printer.Print "Name";
    Printer.CurrentX = 1500: Printer.Print ": " & Star_Nm:
    Printer.Print "-------------------------------------------------------------------"
    Printer.Print "Phone Number";
    Printer.CurrentX = 1500: Printer.Print ": " & Trim(Star_Phone):
'    Printer.Print "":    Printer.Print ""
    Printer.Print "Email";
    Printer.CurrentX = 1500: Printer.Print ": " & Trim(Star_Email):
'    Printer.Print "":    Printer.Print ""
    Printer.Print "-------------------------------------------------------------------"
     Printer.Print "New Phone";
    Printer.CurrentX = 1500: Printer.Print ": "
    Printer.Print ""
    Printer.Print "New Email";
    Printer.CurrentX = 1500: Printer.Print ": "
'    Printer.Print "":    Printer.Print ""
    Printer.Print "-------------------------------------------------------------------"
    Printer.EndDoc
End Sub

Public Sub CetakDataEmail()
    Printer.Font.name = "Printer 17cpi"
    Printer.Font.Size = 9
    Printer.Print "-------------------------------------------------------------------"
    Printer.Print "Name";
    Printer.CurrentX = 1500: Printer.Print ": "
    Printer.Print "-------------------------------------------------------------------"
    Printer.Print "Phone Number";
    Printer.CurrentX = 1500: Printer.Print ": "
'    Printer.Print "":    Printer.Print ""
    Printer.Print "Email";
    Printer.CurrentX = 1500: Printer.Print ": "
'    Printer.Print "":    Printer.Print ""
    Printer.Print "-------------------------------------------------------------------"
    Printer.EndDoc
End Sub
