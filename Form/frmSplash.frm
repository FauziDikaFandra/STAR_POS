VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3300
      Top             =   3300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   3690
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   6765
   End
   Begin VB.Label lblversi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4725
      TabIndex        =   1
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "POINT OF SALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3150
      TabIndex        =   0
      Top             =   2925
      Width           =   3390
   End
   Begin VB.Image Image1 
      Height          =   2940
      Left            =   300
      Picture         =   "frmSplash.frx":08CA
      Stretch         =   -1  'True
      Top             =   225
      Width           =   6315
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Byte

Private Sub Form_Load()
    lblversi = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub Timer1_Timer()
    Call Mulai
End Sub

Private Sub Mulai()
Dim RsAddLinked As New ADODB.Recordset

'    If Not Linked Then
'        StrConLoc = "Provider = SQLOLEDB.1; Persist Security Info=False;" & _
'        "Data Source=" & Environ("computername") & ";" & _
'        "Initial Catalog=" & Cfg_Get("Local", "DatabaseName", App.Path & "\config.ini") & ";" & _
'        "User ID=" & Cfg_Get("Local", "LoginId", App.Path & "\config.ini") & ";" & _
'        "Password=" & decrypt(Cfg_Get("Local", "Password", App.Path & "\config.ini")) & ";"
'    End If
    
    Call OpenKoneksi(ConnLocal, "Local")
'    If ConnLocal.State = False Then

    If Linked Then Call OpenKoneksi(ConnServer, "Server")
    
    RsAddLinked.Open "select srvname from master..sysservers where srvname='" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & "'", ConnLocal, adOpenStatic, adLockReadOnly
        If RsAddLinked.EOF Then ConnLocal.Execute "exec sp_addlinkedserver '" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & "'"
        If RsAddLinked.EOF Then ConnLocal.Execute "EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & "',@useself=N'False',@locallogin=N'sa',@rmtuser=N'sa',@rmtpassword='" & decrypt(Cfg_Get("Local", "Password", App.Path & "\config.ini")) & "'"
    RsAddLinked.Close
    
    RsAddLinked.Open "select srvname from master..sysservers where srvname='" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "'", ConnLocal, adOpenStatic, adLockReadOnly
        If RsAddLinked.EOF And VSvr <> "" Then ConnLocal.Execute "exec sp_addlinkedserver '" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "'"
        If RsAddLinked.EOF Then ConnLocal.Execute "EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "',@useself=N'False',@locallogin=N'sa',@rmtuser=N'sa',@rmtpassword='" & decrypt(Cfg_Get("Server", "Password", App.Path & "\config.ini")) & "'"
    RsAddLinked.Close: Set RsAddLinked = Nothing
    
    If Isi_Parameter Then
        Unload Me
        xx = 0
        Call CEKLPT(True)
        frmLogin.Show
    End If
    
    xx = xx + 1
    If xx = 3 Then
        MsgBox "Tabel Branches tidak lengkap " & vbNewLine & _
        "Harap hubungi IT. ", vbCritical + vbOKOnly, "Oops.."
        End
    End If
End Sub

Private Sub OpenKoneksi(SrvName As ADODB.Connection, Strheader)
On Error GoTo ErrH
    Set SrvName = New ADODB.Connection
    SrvName.CommandTimeout = 3000
    With SrvName
        .CursorLocation = adUseClient
        Select Case Strheader
            Case "Server": .ConnectionString = StrConSvr
            Case "Local": .ConnectionString = StrConLoc
        End Select
        .Open
    End With
    Exit Sub
    
ErrH:
    If Err.Number = "-2147467259" Then
        MsgBox "Database " & Strheader & " tidak terkoneksi. " & vbNewLine & _
                "Harap hubungi IT. ", vbCritical + vbOKOnly, "Oops.."
    Else
        MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    End If
    Call SaveLog("OpenKoneksi " & Err.Description & " " & Err.Number)
    End
End Sub

Private Function Isi_Parameter() As Boolean
On Error GoTo ErrH
Dim Rst As New ADODB.Recordset, tgl_aktif As Date

    Isi_Parameter = True
    Rst.Open "Select * from Branches Where Branch_ID = '" & VBranch_ID & "' ", _
             ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If Rst.EOF Then
        If VPing = "ONLINE" Then
            ConnLocal.Execute "insert into Branches(Branch_ID, Branch_Name, Company_Name, Address1, Address2, City, Zip_Code, Country, Phone," & _
            "Fax, Loyalty_Check, Flag_SOD, Date_Yesterday, Date_Current, Password_Valid, Voucher_Get_Point)" & _
            "(select Branch_id, Branch_Name, Company_Name, Address1, Address2, City, Zip_Code, Country, Phone," & _
            "Fax , Loyalty_Check, 0, Date_Yesterday, Date_Current, Password_Valid, Voucher_Get_Point " & _
            "FROM [" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "]." & _
            Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini") & ".dbo.branches  " & _
            "where branch_id = '" & VBranch_ID & "')"
            
            Isi_Parameter = False
        Else
            MsgBox "Tabel Branches tidak lengkap", vbCritical + vbOKOnly, "Oops.."
            End
        End If
    Else
        Tulis(10) = Rst!branch_name
        Tulis(11) = Rst!company_name
        Tulis(12) = Rst!address1
        Tulis(13) = Rst!address2
        Tulis(14) = Rst!city
        Tulis(15) = Rst!zip_code
        tgl_aktif = Rst!date_current
    End If
        
    Rst.Close
    
    'Penambahan Harga Point dan Path Email By Variable
    If Linked Then
    Rst.Open "Select * from [Parameters] Where [Type] = 'HargaPoint'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Rst.EOF Then
    VHargaPoint = "1000"
    Else
    VHargaPoint = Rst!Value
    End If
    Rst.Close
    
    Rst.Open "Select * from [Parameters] Where [Type] = 'PathEmail'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Rst.EOF Then
    EReceiptEmail = "D:\EReceiptEmail\"
    Else
    EReceiptEmail = Rst!Value
    End If
    Rst.Close
    
    Rst.Open "Select * from [Parameters] Where [Type] = 'Email'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Rst.EOF Then
    StoreEmail = "fauzi.dika@stardeptstore.com"
    Else
    StoreEmail = Rst!Value
    End If
    Rst.Close
    
    Rst.Open "Select * from [Parameters] Where [Type] = 'Loc_Voucher'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Rst.EOF Then
    loc_voucher = "A"
    Else
    loc_voucher = Rst!Value
    End If
    Rst.Close
    
    Else
    VHargaPoint = "1000"
    EReceiptEmail = "D:\EReceiptEmail\"
    StoreEmail = "fauzi.dika@stardeptstore.com"
    End If
    'End
    If PathEmail <> "" Then
    If FolderExists(PathEmail) = True Then
        Dim intFile As Integer
        Dim strFile As String
        Dim strFolderPath As String
        strFile = PathEmail & "\JgnDihapus.txt"
        intFile = FreeFile
        Open strFile For Output As #intFile
            Print #intFile, "Oke"
        Close #intFile
        If Dir(PathEmail & "\BACKUP", vbDirectory) = "" Then
          MkDir PathEmail & "\BACKUP"
        End If
        'Kill PathEmail & "\*.*"
    Else
        MkDir PathEmail & "\"
    End If
    End If
    
        
        

    
    
    
    Rst.Open "Select * from Cash_Register Where Branch_ID = '" & VBranch_ID & "' and Cash_register_Id = '" & _
            VReg_ID & "' ", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
    If Rst.EOF Then
        If VPing = "ONLINE" Then
            ConnLocal.Execute "insert into Cash_Register(Branch_ID, Cash_Register_ID, Spending_Program_ID, Store_Type, Void_Flag, Item_Correct_Flag, Return_Flag, Cancel_Flag, Discount_Flag, Tax_Flag," & _
            "Calc_Point_Flag, Bill_No, Shift, Reset_Date, Last_Reset_Date, Disc_1, Disc_2, Disc_3, Disc_4, Disc_5, Disc_6, Disc_7, Footer_1, Footer_2, Footer_3, Footer_4, Footer_5, Footer_6, Active_Status, ZReset_Status, NPWP, SMessage1, SMessage2)" & _
            "(select Branch_ID, Cash_Register_ID, Spending_Program_ID, Store_Type, Void_Flag, Item_Correct_Flag, Return_Flag, Cancel_Flag, Discount_Flag, Tax_Flag," & _
            "Calc_Point_Flag, Bill_No, Shift, Reset_Date, Last_Reset_Date, Disc_1, Disc_2, Disc_3, Disc_4, Disc_5, Disc_6, Disc_7, Footer_1, Footer_2, Footer_3, Footer_4, Footer_5, Footer_6, Active_Status, ZReset_Status, NPWP, SMessage1, SMessage2 " & _
            "FROM [" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "]." & _
            Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini") & ".dbo.cash_register  " & _
            "where branch_id = '" & VBranch_ID & "' and Cash_Register_ID='" & VReg_ID & " ')"
            
            Isi_Parameter = False
        Else
            MsgBox "Tabel Cash Register tidak lengkap", vbCritical + vbOKOnly, "Oops.."
            End
        End If
    Else
        If Format(Rst!reset_date, "YYYY-MM-DD") > Format(GetSrvDate, "YYYY-MM-DD") Then
            If Format(tgl_aktif, "YYYY-MM-DD") > Format(GetSrvDate, "YYYY-MM-DD") Then
                MsgBox "Transaksi hari ini sudah closed (Z-Reset)", vbCritical + vbOKOnly, "Oops.."
                End
            End If
        End If
        
        VShift = Rst!Shift
    
        Tulis(1) = Rst!Footer_1
        Tulis(2) = Rst!Footer_2
        Tulis(3) = Rst!Footer_3
        Tulis(4) = Rst!Footer_4
        Tulis(5) = Rst!Footer_5
        Tulis(6) = Rst!Footer_6
        Tulis(7) = Rst!SMessage1
        Tulis(8) = Rst!SMessage2
        Tulis(9) = Rst!NPWP
    End If
    Rst.Close: Set Rst = Nothing
    Exit Function

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Isi_Parameter " & Err.Description & " " & Err.Number)
    End
End Function

Public Function CEKLPT(cek As Boolean)
Dim strOut, strPRT As String
 
    If PPort = "lpt1" Then
        strPRT = 379
        Do While cek <> False
             strOut = Str(Inp(Val("&H" + strPRT)))
             Out Val("&H" + strPRT), Val(strOut)
             If strPRT = 379 And strOut = 127 Then
                 MsgBox "Printer belum dinyalakan", vbExclamation, "Warning"
             Else
                If strPRT = 379 And strOut = 119 Or strOut = 103 Then
                    MsgBox "Kertas Printer Habis", vbExclamation, "Warning"
                Else
                    cek = False
                End If
             End If
         Loop
    End If
End Function
