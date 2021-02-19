Attribute VB_Name = "Modul"
Option Explicit
Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Function Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const cMAXLEN = 255 'English name of country
Private Const LOCALE_SENGCOUNTRY = &H1002
Private Const LOCALE_SYSTEM_DEFAULT& = &H800
Private Const LOCALE_USER_DEFAULT& = &H400
Private Declare Function apiGetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private PingIPv4 As PingIPv4
Public StrukEmail As Boolean
Public ConnServer As New ADODB.Connection, ConnLocal As New ADODB.Connection
Public StrConSvr As String, StrConLoc As String
Public KeyStroke(27 To 121) As String, Tulis(0 To 20) As String
Public InputEmail, InputNama, InputNoTlp As String

Public VBranch_ID As String, VReg_ID As String, VSvr As String
Public DiscStarProID As Integer, isLimitCC As Integer, VDiscBySTAR As Long, Tipe3Total As Long, LimitTipe3 As Long
Public CDCom As String, CDSet As String, PPort As String, Log_P As String, TipKom As Byte, VGaris As String

Public VKasir_ID As String, VKasir_Nm As String, VShift As Byte, OnlyCheckPromo As Integer, isECR As Integer, ECRComm As Integer, AcaraMKT As Integer
Public Star_Nm As String, Star_Pt As Integer, Star_Id As String, Star_Freq As String, Star_Omz As String, Star_No As String, Star_Ext1 As String, Exp_Point As String, Expired_Date As String, Star_Gender As String
Public Star_Phone As String, Star_Email As String, Star_updsts As Byte, Footer6str As Long

Public StrSQL As String, VNomor As String, VPing As String, VSuper_Nm As String, VKary As String, VCekKartu As String, VHargaPoint As String, EReceiptEmail As String, StoreEmail As String, PathEmail  As String, loc_voucher  As String
Public VOK As Boolean, VCopen As Boolean, VAda_Promo As Boolean, VTanya As Boolean, VIsSSC As Boolean, VIsKKG As Boolean, MSCTlp As Boolean, UpdateStatusSeq As Boolean, UpdateStatusSeqDetail As Boolean, ScanApps As Boolean
Public VBonus_Point As Byte
Public SeqCount(10) As String
Public SeqCountInt As Integer

Sub Main()
Dim RegSetting As String

    If App.PrevInstance = True Then End
    
    RegSetting = fLocaleInfo(LOCALE_SENGCOUNTRY)
    If RegSetting <> "United States" Then
        MsgBox "Anda menggunakan Regional Setting " & RegSetting & vbCrLf & _
        "Anda harus mengganti Regional Setting menjadi English(United States)", 64, "Oops.."
        End
    End If
    
    Set PingIPv4 = New PingIPv4
    
    StrConSvr = "Provider = SQLOLEDB.1; Persist Security Info=False;" & _
            "Data Source=" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & ";" & _
            "Initial Catalog=" & Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini") & ";" & _
            "User ID=" & Cfg_Get("Server", "LoginId", App.Path & "\config.ini") & ";" & _
            "Password=" & decrypt(Cfg_Get("Server", "Password", App.Path & "\config.ini")) & ";"
    
    StrConLoc = "Provider = SQLOLEDB.1; Persist Security Info=False;" & _
            "Data Source=" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & ";" & _
            "Initial Catalog=" & Cfg_Get("Local", "DatabaseName", App.Path & "\config.ini") & ";" & _
            "User ID=" & Cfg_Get("Local", "LoginId", App.Path & "\config.ini") & ";" & _
            "Password=" & decrypt(Cfg_Get("Local", "Password", App.Path & "\config.ini")) & ";"

    VBranch_ID = Cfg_Get("RegisterInfo", "BranchID", App.Path & "\Config.ini")
    VReg_ID = Cfg_Get("RegisterInfo", "RegID", App.Path & "\Config.ini")
    VSvr = Cfg_Get("Server", "ServerName", App.Path & "\config.ini")
    
    CDCom = Cfg_Get("Device", "CD_Com", App.Path & "\config.ini")
    CDSet = Cfg_Get("Device", "CD_Set", App.Path & "\config.ini")
    PPort = Cfg_Get("Device", "PrinterPort", App.Path & "\config.ini")
    Log_P = Cfg_Get("Device", "Log_Path", App.Path & "\config.ini")
    TipKom = Cfg_Get("Device", "Touch", App.Path & "\config.ini")
    isECR = Cfg_Get("Device", "isECR", App.Path & "\config.ini")
    ECRComm = Cfg_Get("Device", "ECRComm", App.Path & "\config.ini")
    AcaraMKT = Cfg_Get("Device", "AcaraMKT", App.Path & "\config.ini")
    PathEmail = Cfg_Get("Device", "Email_Path", App.Path & "\config.ini")
    
    VPing = IIf(VSvr = "", "OFFLINE", "ONLINE")
    If TipKom = 1 Then
        VGaris = "----------------------------------------------------------------------"
    Else
        VGaris = "------------------------------------------"
    End If
    RegSetting = Cfg_Get("Apl", "Revision", "\\" & Cfg_Get("Server", "ServerName", App.Path & "\config.ini") & "\Apl\AplCfg.ini")
    
    If RegSetting <> "" Then
        If RegSetting > App.Revision Then
            MsgBox "Aplikasi belum diupdate" & vbNewLine & "Harap hubungi IT.", vbCritical + vbOKOnly, "Oops.."
            Exit Sub
        End If
    End If
    
    If VBranch_ID <> "" And VReg_ID <> "" Then
        frmSplash.Show
    Else
        MsgBox "File konfigurasi belum lengkap", vbCritical + vbOKOnly, "Oops.."
    End If
End Sub

Public Function decrypt(ByVal unpass As String) As String
Dim x As Integer
Dim awal As String, kembali As String
    
    x = 1
    awal = ""
    
    Do While x <= Len(Trim(unpass))
       kembali = Mid(unpass, x, 3)
       x = x + 3
       awal = awal + Chr((Val(kembali) + 11) / 3 - 5)
    Loop
    decrypt = awal
End Function

Function isValidEmail(myEmail As String) As Boolean
' existence of '@'
If occurrenceOf(myEmail, "@") = 0 Then
isValidEmail = False
Exit Function
End If
' existence of '@' more than once.
If occurrenceOf(myEmail, "@") > 1 Then
isValidEmail = False
Exit Function
End If
' existence of '.'(dot).
If occurrenceOf(myEmail, ".") = 0 Then
isValidEmail = False
Exit Function
End If
' existence of space
If occurrenceOf(myEmail, " ") <> 0 Then
isValidEmail = False
Exit Function
End If
' the first char is digit
If Left$(myEmail, 1) Like "[0-9]" Then
isValidEmail = False
Exit Function
End If
' existence of . in the ID part, or x@.x pattern
If (InStr(1, myEmail, "@") + 1) >= InStrRev(myEmail, ".") Then
isValidEmail = False
Exit Function
End If
isValidEmail = True
End Function

Function occurrenceOf(Source As String, char As String)
Dim I As Integer, j As Integer
Dim myCount As Integer
myCount = 0
I = InStr(1, Source, char)
Do While I > 0
myCount = myCount + 1
I = InStr(I + 1, Source, char)
Loop
occurrenceOf = myCount
End Function

Private Function fLocaleInfo(lngLCType As Long) As String
Dim strLCData As String, lngData As Long
Dim lngX As Long

  strLCData = String$(cMAXLEN, 0)
  lngData = cMAXLEN - 1
  lngX = apiGetLocaleInfo(LOCALE_USER_DEFAULT, lngLCType, strLCData, lngData)
  If lngX <> 0 Then
      fLocaleInfo = Left$(strLCData, lngX - 1)
  End If
End Function

Public Function Linked() As Boolean
Dim ResolveResult As RESOLVE_ERRORS
Dim IP As String

    If VPing = "OFFLINE" Then
        Linked = False
    Else
        ResolveResult = PingIPv4.Resolve(VSvr, IP)
        If ResolveResult = RES_SUCCESS Then
            If PingIPv4.Ping(IP) Then
                Linked = True
            Else
                Linked = False
                MsgBox "Status : OFFLINE", vbInformation + vbOKOnly, "Oops.."
                Call SaveLog("OFFLINE")
                VPing = "OFFLINE"
            End If
        Else
            MsgBox "Server Name/IP tidak valid", vbCritical + vbOKOnly, "Oops.."
            End
        End If
    End If
End Function

Public Function FolderExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(sFullPath)
End Function

Public Function Cfg_Get(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
Const cparmLen = 50
Dim sReturn As String * cparmLen
Dim sDefault As String * cparmLen
Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, sKey _
            , sDefault, sReturn, cparmLen, sINIFileName)
    Cfg_Get = Mid(sReturn, 1, lLength)
End Function

' Note: Add a reference to the Microsoft CDO library.
Public Function SendEmail(ByVal strSender As String, _
                        ByVal strRecipient As String, _
                        Optional ByVal attachment As String, _
                        Optional ByVal strCc As String, _
                        Optional ByVal strBcc As String _
                         ) As Boolean
    Dim cdoMsg As New CDO.Message
    Dim cdoConf As New CDO.Configuration
    Dim schema As String
    Dim Flds
    Dim strHTML
    
    On Error GoTo ErrTrap
    Const cdoSendUsingPort = 2
    
    'Set cdoMsg =  CreateObject("CDO.Message")
    'Set cdoConf = CreateObject("CDO.Configuration")
    
    Set Flds = cdoConf.Fields
        
    schema = "http://schemas.microsoft.com/cdo/configuration/"

    With Flds
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = "smtp.gmail.com"
        .Item(schema & "smtpserverport") = 465
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = StoreEmail
        .Item(schema & "sendpassword") = "31Jul2017"
        .Item(schema & "smtpusessl") = 1
        .Update
    End With
    
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = strRecipient
        .From = strSender
        .Subject = "E-Receipt " & attachment
        .TextBody = " Kepada Yth. Pelanggan STAR DEPARTMENT STORE , " & vbNewLine & vbNewLine & " " & _
                    "Terima Kasih telah berbelanja di toko kami. " & vbNewLine & vbNewLine & " " & _
                    "Terlampir e-receipt pembelanjaan Anda di toko STAR DEPARTMENT STORE. " & vbNewLine & vbNewLine & " " & _
                    "Hormat Kami, " & vbNewLine & vbNewLine & " " & _
                    "PT STAR MAJU SENTOSA"
        .AddAttachment (PathEmail & "\" & attachment & ".pdf")
        If strCc <> "" Then .cc = strCc
        If strBcc <> "" Then .BCC = strBcc
        .Send
    End With
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
        
    SendEmail = True
    Exit Function
ErrTrap:
'Err.Raise Err.Number, "", "Error from Functions.SendEmail" & Err.Description
    SendEmail = False
End Function

Public Function GetSrvDate() As Date
On Error GoTo ErrH
Dim RsTglVsvr As New ADODB.Recordset

    If Linked Then
        RsTglVsvr.Open "select getdate() as srvdt", ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsTglVsvr.Open "select getdate() as srvdt", ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
    
    GetSrvDate = RsTglVsvr!srvdt
    RsTglVsvr.Close: Set RsTglVsvr = Nothing
    Exit Function

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog("GetSrvDate " & Err.Description & " " & Err.Number)
    GetSrvDate = Now()
End Function

Public Function Super(lvl As Byte) As Boolean
    VOK = False
    VSuper_Nm = ""
    frmValid.VLevelApp = lvl
    frmValid.Show 1
    DoEvents
    Super = VOK
End Function

Public Sub SQLQuery(kueri As String)
On Error GoTo ErrH
Dim Flg As Integer
Flg = 0
'penambahan rollback point
1:
    ConnLocal.BeginTrans
    ConnLocal.Execute kueri

    If Linked Then
        If Flg = 1 Then
            ConnServer.RollbackTrans
        End If
On Error GoTo ErrD
    ConnServer.BeginTrans
    ConnServer.Execute kueri
    ConnServer.CommitTrans
GoTo 2
ErrD:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog("SQLQueryServer " & kueri & " " & Err.Description & " " & Err.Number)
    ConnServer.RollbackTrans
2:
    End If
    
    ConnLocal.CommitTrans
    Exit Sub

ErrH:
    ConnLocal.RollbackTrans
    Flg = Flg + 1
    If Flg <= 1 Then
        GoTo 1
    Else
        MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog("SQLQueryLocal " & kueri & " " & Err.Description & " " & Err.Number)
End If

    
   
End Sub

Public Sub SaveLog(ByVal Msg1 As String)
On Error GoTo ErrH

    If Log_P <> "" Then
        Open Log_P & "\LOG_" & VReg_ID & "_" & Format(Now(), "mmddyy") & ".txt" For Append As #2
        Print #2, Now() & " " & Msg1
        Close #2
    End If
    Exit Sub
    
ErrH:
   MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
   Log_P = ""
End Sub

Public Sub SendMail(ByVal EmailTo As String)
On Error GoTo ErrH

    'Dim oSmtp As New EASendMailObjLib.Mail
    'oSmtp.LicenseCode = "TryIt"

    'oSmtp.FromAddr = "gmailid@gmail.com"

    'oSmtp.AddRecipientEx "support@emailarchitect.net", 0

    'oSmtp.Subject = "test email struck"

    'oSmtp.BodyText = "this is email from POS struck"

    'oSmtp.ServerAddr = "smtp.gmail.com"

    '' set 25 or 587 port
    'oSmtp.ServerPort = 587

    'oSmtp.SSL_init

    'oSmtp.UserName = "gmailid@gmail.com"
    'oSmtp.Password = "yourpassword"

    ''MsgBox "start to send email ..."

    'If oSmtp.SendMail() = 0 Then
        'MsgBox "email was sent successfully!"
    ''Else
        'MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    'End If
    'Exit Sub
    
ErrH:
   MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
   Log_P = ""
End Sub

Public Sub CDisplay(tex1 As String, Tex2 As String)
On Error Resume Next
Dim ConCom As New MSComm
    
    If CDCom = "" Then Exit Sub
    With ConCom
        .CommPort = CDCom
        .Settings = CDSet
        .InputMode = comInputModeText

        tex1 = Space((20 - Len(tex1)) / 2) & tex1
        Tex2 = Space((20 - Len(Tex2)) / 2) & Tex2
'----------------------Normal-----------------------------
        If .PortOpen = False Then
            .PortOpen = True
            .Output = Chr(27) & "[2J" 'bersihkan display
            .Output = Chr(27) & "[" & Chr(&H31 + 0) & ";" & Chr(&H31 + 0) & "H" & tex1
            .Output = Chr(27) & "[" & Chr(&H31 + 1) & ";" & Chr(&H31 + 0) & "H" & Tex2
            .PortOpen = False
        End If
'---------------------POSIFLEX----------------------------
'        If .PortOpen = False Then
'            .PortOpen = True
'             .Output = Chr(12) 'bersihkan display
'             .Output = Chr(27) & Chr(81) & Tex1 & vbCrLf
'             .Output = Chr(27) & Chr(81) & Tex2
'           .PortOpen = False
'        End If

    End With
End Sub

Public Function UbahChar(Kata As String) As String
    UbahChar = Replace(Kata, "'", "''")
End Function

'Private Function encrypt(ByVal pass As String) As String
'Dim ubah As String
'Dim hit As Integer, Y As Integer
'
'    Y = 1
'    ubah = ""
'
'    Do While Y <= Len(Trim(pass))
'       hit = (Asc(Mid(pass, Y, 1)) + 5) * 3 - 11
'       ubah = ubah + LTrim(hit)
'       Y = Y + 1
'    Loop
'    encrypt = ubah
'End Function
