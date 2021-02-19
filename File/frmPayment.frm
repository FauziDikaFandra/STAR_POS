VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9540
   ControlBox      =   0   'False
   Icon            =   "frmPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   1965
      Index           =   1
      Left            =   4125
      TabIndex        =   4
      Top             =   1800
      Width           =   5265
      Begin TDBText6Ctl.TDBText txtno_kartu 
         Height          =   390
         Left            =   300
         TabIndex        =   5
         Top             =   300
         Width           =   4665
         _Version        =   65536
         _ExtentX        =   8229
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0936
         Key             =   "frmPayment.frx":0954
         BackColor       =   -2147483643
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   300
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
      Begin TDBText6Ctl.TDBText txtnama 
         Height          =   390
         Left            =   300
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   825
         Width           =   4665
         _Version        =   65536
         _ExtentX        =   8229
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":0998
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":09FC
         Key             =   "frmPayment.frx":0A1A
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   100
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
      Begin TDBNumber6Ctl.TDBNumber txtcredit 
         Height          =   390
         Left            =   300
         TabIndex        =   12
         Top             =   1350
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":0A6C
         Caption         =   "frmPayment.frx":0A8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0AF2
         Keys            =   "frmPayment.frx":0B10
         Spin            =   "frmPayment.frx":0B5A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1994981377
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
   End
   Begin VB.CommandButton cmdpay 
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
      Index           =   1
      Left            =   4200
      Picture         =   "frmPayment.frx":0B82
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4575
      Width           =   1140
   End
   Begin VB.CommandButton cmdpay 
      Caption         =   "CLOSE"
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
      Index           =   0
      Left            =   8250
      Picture         =   "frmPayment.frx":1584
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4575
      Width           =   1140
   End
   Begin VB.CommandButton cmdpay 
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
      Index           =   2
      Left            =   5475
      Picture         =   "frmPayment.frx":1F86
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4575
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   4125
      TabIndex        =   13
      Top             =   150
      Width           =   5265
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "999,999,999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   540
         Left            =   2250
         TabIndex        =   19
         Top             =   75
         Width           =   2790
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "TOTAL : Rp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   540
         Left            =   75
         TabIndex        =   18
         Top             =   75
         Width           =   2115
      End
      Begin VB.Label lblbayar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "999,999,999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   2250
         TabIndex        =   17
         Top             =   600
         Width           =   2790
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "BAYAR    : Rp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   75
         TabIndex        =   16
         Top             =   600
         Width           =   2115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         X1              =   2625
         X2              =   5100
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label lblsisa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "999,999,999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   2250
         TabIndex        =   15
         Top             =   1125
         Width           =   2790
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "SISA        : Rp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   75
         TabIndex        =   14
         Top             =   1125
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Height          =   165
         Left            =   0
         TabIndex        =   20
         Top             =   975
         Width           =   5190
      End
   End
   Begin TrueOleDBGrid80.TDBGrid GridPaid 
      Bindings        =   "frmPayment.frx":2988
      Height          =   2040
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5550
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   3598
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No"
      Columns(0).DataField=   "Seq"
      Columns(0).DataWidth=   11
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tipe"
      Columns(1).DataField=   "Description"
      Columns(1).DataWidth=   30
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nilai"
      Columns(2).DataField=   "Paid_Amount"
      Columns(2).DataWidth=   22
      Columns(2).NumberFormat=   "#,##0"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No Kartu"
      Columns(3).DataField=   "Credit_Card_No"
      Columns(3).DataWidth=   16
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Kartu"
      Columns(4).DataField=   "Credit_Card_Name"
      Columns(4).DataWidth=   50
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3228"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3122"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1905"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(2)._AlignLeft=0"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=3493"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3387"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=5159"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=5054"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid GridPay_Types 
      Bindings        =   "frmPayment.frx":299D
      Height          =   4740
      Left            =   150
      TabIndex        =   0
      Top             =   675
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   8361
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Payment_Types"
      Columns(0).DataField=   "Payment_Types"
      Columns(0).DataWidth=   4
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "TIPE PEMBAYARAN"
      Columns(1).DataField=   "Description"
      Columns(1).DataWidth=   30
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Types"
      Columns(2).DataField=   "Types"
      Columns(2).DataWidth=   2
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Seq"
      Columns(3).DataField=   "Seq"
      Columns(3).DataWidth=   11
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5080"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4974"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1005"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=900"
      Splits(0)._ColumnProps(13)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1561"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1455"
      Splits(0)._ColumnProps(18)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(3)._AlignLeft=0"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Tdata1 
      Height          =   330
      Left            =   180
      Top             =   5610
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Tdata2 
      Height          =   330
      Left            =   180
      Top             =   5970
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TDBNumber6Ctl.TDBNumber vpay 
      Height          =   315
      Left            =   1500
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   556
      Calculator      =   "frmPayment.frx":29B2
      Caption         =   "frmPayment.frx":29D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayment.frx":2A3E
      Keys            =   "frmPayment.frx":2A5C
      Spin            =   "frmPayment.frx":2AA6
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   -999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   113967105
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   1440
      Index           =   0
      Left            =   4125
      TabIndex        =   2
      Top             =   1800
      Width           =   5265
      Begin TDBNumber6Ctl.TDBNumber txtcash 
         Height          =   390
         Left            =   225
         TabIndex        =   10
         Top             =   225
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":2ACE
         Caption         =   "frmPayment.frx":2AEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2B52
         Keys            =   "frmPayment.frx":2B70
         Spin            =   "frmPayment.frx":2BBA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin TDBNumber6Ctl.TDBNumber txtkembali 
         Height          =   390
         Left            =   225
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   750
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":2BE2
         Caption         =   "frmPayment.frx":2C02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2C6C
         Keys            =   "frmPayment.frx":2C8A
         Spin            =   "frmPayment.frx":2CD4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
   End
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   1440
      Index           =   2
      Left            =   4125
      TabIndex        =   7
      Top             =   1800
      Width           =   5265
      Begin TDBText6Ctl.TDBText txtno_voc 
         Height          =   390
         Left            =   300
         TabIndex        =   8
         Top             =   300
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":2CFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2D6C
         Key             =   "frmPayment.frx":2D8A
         BackColor       =   -2147483643
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   11
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
      Begin TDBNumber6Ctl.TDBNumber txtvoucher 
         Height          =   390
         Left            =   300
         TabIndex        =   25
         Top             =   840
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":2DCE
         Caption         =   "frmPayment.frx":2DEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2E54
         Keys            =   "frmPayment.frx":2E72
         Spin            =   "frmPayment.frx":2EBC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   12255233
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
   End
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   2115
      Index           =   3
      Left            =   4125
      TabIndex        =   27
      Top             =   1800
      Width           =   5265
      Begin VB.TextBox txtharga_point 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "800"
         Top             =   1050
         Width           =   840
      End
      Begin TDBNumber6Ctl.TDBNumber txttukar_point 
         Height          =   390
         Left            =   150
         TabIndex        =   28
         Top             =   1350
         Width           =   2190
         _Version        =   65536
         _ExtentX        =   3863
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":2EE4
         Caption         =   "frmPayment.frx":2F04
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2F80
         Keys            =   "frmPayment.frx":2F9E
         Spin            =   "frmPayment.frx":2FE8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   112918529
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin TDBNumber6Ctl.TDBNumber txtsaldo_point 
         Height          =   390
         Left            =   150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   825
         Width           =   2190
         _Version        =   65536
         _ExtentX        =   3863
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":3010
         Caption         =   "frmPayment.frx":3030
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":30A2
         Keys            =   "frmPayment.frx":30C0
         Spin            =   "frmPayment.frx":310A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12632256
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1967521797
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin TDBText6Ctl.TDBText txtcard_no 
         Height          =   390
         Left            =   150
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   300
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":3132
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":319E
         Key             =   "frmPayment.frx":31BC
         BackColor       =   12632256
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   200
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
      Begin TDBNumber6Ctl.TDBNumber txtpoint 
         Height          =   390
         Left            =   3825
         TabIndex        =   34
         Top             =   1350
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2275
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":320E
         Caption         =   "frmPayment.frx":322E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":328A
         Keys            =   "frmPayment.frx":32A8
         Spin            =   "frmPayment.frx":32F2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12632256
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1996619777
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin TDBNumber6Ctl.TDBNumber txtnilai_point 
         Height          =   390
         Left            =   3825
         TabIndex        =   35
         Top             =   825
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2275
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":331A
         Caption         =   "frmPayment.frx":333A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":3396
         Keys            =   "frmPayment.frx":33B4
         Spin            =   "frmPayment.frx":33FE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12632256
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,##0;(#,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,##0;(#,###,###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -9999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   32
         Top             =   1125
         Width           =   315
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   30
         Top             =   1125
         Width           =   315
      End
   End
   Begin VB.Label vstatus 
      Height          =   315
      Left            =   2775
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label vno_trans 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   5625
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "PAYMENT TYPES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   3840
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    lbltotal = Format(vpay, "#,##0")
    lblbayar = 0
    lblsisa = Format(vpay, "#,##0")
End Sub

Private Sub Form_Load()
    Tdata1.ConnectionString = StrConLoc
    Tdata2.ConnectionString = StrConLoc
    vno_trans = VNomor

    Tdata1.RecordSource = "SELECT Payment_Types, Description, Types, Seq From Payment_Types where Seq<30 ORDER BY Seq"
    Tdata2.RecordSource = "SELECT aa.Seq, bb.Description, Paid_Amount, Credit_Card_No, Credit_Card_Name " & _
                          "FROM Paid aa INNER JOIN Payment_Types bb ON aa.Payment_Types = bb.Payment_Types" & _
                          " where transaction_number = '" & vno_trans & "' order by aa.seq"

    Tdata1.Refresh
    Tdata2.Refresh
End Sub

Private Sub cmdpay_Click(Index As Integer)
    Select Case Index
    Case 0
        Call GridPay_Types_KeyDown(27, 0)
    Case 1
        GridPay_Types.MovePrevious
        If GridPay_Types.BOF Then GridPay_Types.MoveFirst
    Case 2
        GridPay_Types.MoveNext
        If GridPay_Types.EOF Then GridPay_Types.MoveLast
    End Select
End Sub

Private Sub GridPay_Types_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Select Case GridPay_Types.Columns(2)
            Case "CS"
                txtcash = vpay
                txtcash.SetFocus
            Case "CC", "DC"
                txtcredit = vpay
                txtno_kartu = ""
                txtnama = ""
                txtno_kartu.SetFocus
            Case "SV", "OV"
                txtno_voc.SetFocus
            Case "PR"
                txttukar_point.SetFocus
        End Select
        DoEvents
        SendKeys "{home}+{end}"
    Case 27
        Dim RsHapus As New ADODB.Recordset
        
        If Not Tdata2.Recordset.EOF Then Tdata2.Recordset.MoveFirst
        While Not Tdata2.Recordset.EOF
            If Tdata2.Recordset!Description = "POINT REWARD" Then
                Call SQLQuery("delete from cust_prize_trans where trans_nr='" & Tdata2.Recordset!credit_card_name & "'")
                
                Call SQLQuery("Update card set card_point=card_point + " & Tdata2.Recordset!paid_amount / 800 & "where card_nr = '" & Tdata2.Recordset!credit_card_no & "'")
            End If
            Tdata2.Recordset.MoveNext
        Wend
        Call MySTAR(txtcard_no)
        Call tampil_point
        
        Call SQLQuery("delete from Paid where Transaction_Number = '" & vno_trans & "'")
    
    VNomor = vno_trans
    Unload Me
    End Select
End Sub

Private Sub GridPay_Types_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim x As Byte
    lblmsg.Caption = GridPay_Types.Columns(1)
    
    For x = 0 To 3
        frmpay(x).Visible = False
    Next x
    
    Select Case GridPay_Types.Columns(2)
        Case "CS"
            frmpay(0).Visible = True
        Case "CC", "DC"
            frmpay(1).Visible = True
        Case "SV", "OV"
            frmpay(2).Visible = True
        Case "PR"
            Call tampil_point
            frmpay(3).Visible = True
    End Select
End Sub

Private Sub txtcash_Change()
    txtkembali = Val(txtcash) - vpay
End Sub

Private Sub txtcash_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            GridPay_Types.SetFocus
            If Simpan_Detail(IIf(txtkembali > 0, txtcash - txtkembali, txtcash), "", "") Then Call Cek_Lunas
        Case 27
            GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then KeyAscii = 0    '45 -
End Sub

Private Sub txtno_kartu_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Select Case GridPay_Types.Columns(2)
        Case "CC", "DC"
            If txtno_kartu.Text = "" Then
                txtno_kartu.SetFocus
                Exit Sub
            End If
            
            Dim Kartu As String
            Kartu = LTrim(txtno_kartu.Text)
            
            If Len(Kartu) = 126 Or Len(Kartu) = 121 Or Len(Kartu) = 90 Then
                txtno_kartu = LTrim(Mid(Kartu, 5, 16))
                txtnama = LTrim(Mid(Kartu, 22, 24))
            ElseIf Len(Kartu) = 128 Or Len(Kartu) = 101 Or Len(Kartu) = 210 Or Len(Kartu) = 233 Then
                txtno_kartu = LTrim(Mid(Kartu, 5, 16))
                txtnama = LTrim(Mid(Kartu, 22, 26))
            ElseIf Len(Kartu) = 103 Then
                txtno_kartu = LTrim(Mid(Kartu, 5, 17))
                txtnama = LTrim(Mid(Kartu, 23, 26))
            ElseIf Len(Kartu) = 51 Then
                txtno_kartu = LTrim(Mid(Kartu, 8, 16))
            Else
                If txtno_kartu <> "" Then Call SaveLog(Kartu & " - KARTU BARU")
                txtno_kartu = LTrim(Mid(Kartu, 5, 16))
            End If
        End Select
        
        txtcredit.SetFocus
        DoEvents
        SendKeys "{home}+{end}"
    Case 27
        GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtcredit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If txtcredit <= vpay Then
                GridPay_Types.SetFocus
                If Simpan_Detail(txtcredit, Left(txtno_kartu, 16), Left(txtnama, 50)) Then
                    txtno_kartu = ""
                    txtnama = ""
                    txtcredit = 0
                    Call Cek_Lunas
                End If
            Else
                MsgBox "Pembayaran dengan kartu kredit/debit tidak boleh " & vbNewLine & _
                "melebihi sisa yang harus dibayar", vbOKOnly + vbInformation, "Oops.."
                txtcredit.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            End If
        Case 27
            GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtno_voc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RsDobel As New ADODB.Recordset

    Select Case KeyCode
    Case 13
        Select Case GridPay_Types.Columns(2)
        Case "SV"
            
            If Linked Then
                txtvoucher = UCase(Cek_Voc(txtno_voc))
                If txtvoucher = 0 Then
                    txtno_voc.SetFocus
                    DoEvents
                    SendKeys "{home}+{end}"
                    Exit Sub
                Else
                    txtvoucher.ReadOnly = True
                    txtvoucher.SetFocus
                End If
            Else
                txtvoucher.ReadOnly = False
                txtvoucher.SetFocus
            End If
        
            RsDobel.Open "select credit_card_no from paid where credit_card_no = '" & txtno_voc & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            If Not RsDobel.EOF Then
                MsgBox "Nomor Voucher Sudah Pernah Dipakai", vbOKOnly + vbInformation, "Oops.."
                txtno_voc.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            End If
            RsDobel.Close:   Set RsDobel = Nothing
            
        Case "OV"
            txtvoucher.ReadOnly = False
            txtvoucher.SetFocus
        End Select
    Case 27
        GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtvoucher_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        If txtvoucher > 0 Then
            GridPay_Types.SetFocus
            If Simpan_Detail(txtvoucher, txtno_voc, "") Then
                txtno_voc = ""
                txtvoucher = 0
                Call Cek_Lunas
            End If
        End If
    Case 27
        GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txttukar_point_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        If txttukar_point.Value > txtsaldo_point Then
            MsgBox "Saldo point tidak mencukupi", vbOKOnly + vbInformation, "Oops.."
            Exit Sub
        End If
        If txttukar_point.Value > 0 Then
            If txtpoint.Value > vpay Then
                MsgBox "Pembayaran dengan point reward tidak boleh " & vbNewLine & _
                "melebihi sisa yang harus dibayar", vbOKOnly + vbInformation, "Oops.."
                txttukar_point.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            Else
                GridPay_Types.SetFocus
                
                If Simpan_Detail(txtpoint, txtcard_no, Pay_Point(txttukar_point, txtcard_no)) Then
                'If Simpan_Detail(txtpoint, Pay_Point(txttukar_point, txtcard_no), "MYSTAR CARD POINT (" & txttukar_point & ")") Then
                    Call tampil_point
                    txttukar_point = 0
                    Call Cek_Lunas
                End If
            End If
        End If
    Case 27
        GridPay_Types.SetFocus
    End Select
End Sub

Private Function Gen_Seq() As String
Dim RsCari As New ADODB.Recordset
    
    RsCari.Open "select MAX(seq) as urut from paid where Transaction_Number = '" & vno_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Gen_Seq = IIf(Not IsNull(RsCari!urut), RsCari!urut + 1, 1)
    RsCari.Close:   Set RsCari = Nothing
End Function

Private Function Cek_Voc(nomor As String) As Long
Dim RsCari As New ADODB.Recordset
Dim RsDobel As New ADODB.Recordset

    Cek_Voc = 0
    RsCari.Open "select v_amt from newvoc where v_no = '" & nomor & _
                "' AND (V_FLAG IS NULL) AND (V_SELL IS NOT NULL)", ConnServer, adOpenForwardOnly, adLockReadOnly
    Cek_Voc = IIf((RsCari.EOF), 0, RsCari!v_amt)
    If Cek_Voc = 0 Then
        MsgBox "Nomor Voucher tidak valid", vbOKOnly + vbInformation, "Oops.."
        RsCari.Close:   Set RsCari = Nothing
        Exit Function
    End If
    RsCari.Close:   Set RsCari = Nothing

    RsDobel.Open "select credit_card_no from paid where credit_card_no = '" & nomor & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Not RsDobel.EOF Then
        MsgBox "Nomor Voucher Sudah Pernah Dipakai", vbOKOnly + vbInformation, "Oops.."
        Cek_Voc = 0
    End If
    RsDobel.Close:   Set RsDobel = Nothing
End Function

Private Function Simpan_Detail(bayar As Double, card_no As String, card_num As String) As Boolean
On Error GoTo ErrH

    Simpan_Detail = False
    
    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No, " & _
            "Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & vno_trans & "','" & GridPay_Types.Columns(0) & "','" & Gen_Seq & _
            "','IDR','1','" & card_no & "','" & UbahChar(card_num) & "'," & bayar & ",'" & VShift & "')"
    
    Tdata2.RecordSource = "SELECT aa.Seq, bb.Description, Paid_Amount, Credit_Card_No, Credit_Card_Name " & _
                        "FROM Paid aa INNER JOIN Payment_Types bb ON aa.Payment_Types = bb.Payment_Types" & _
                        " where transaction_number = '" & vno_trans & "' order by seq"
    Tdata2.Refresh
    GridPaid.MoveLast
    
    vpay = vpay - bayar
    lblbayar = Format(lbltotal - vpay, "#,##0")
    lblsisa = Format(vpay, "#,##0")
    
    Simpan_Detail = True
    Exit Function
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Simpan_Detail " & Err.Description & " " & Err.Number)
End Function

Private Sub Cek_Lunas()
On Error GoTo ErrH
Dim x As Byte

    If vpay <= 0 And vstatus = "SALES" Then 'jika lunas
        Call Paid_To_Sales(vno_trans, lbltotal, IIf(txtkembali > 0, txtkembali, 0))
        GoTo lanjut
    End If
    
    If vpay >= 0 And vstatus = "REFUND" Then 'jika lunas
        Call Paid_To_Sales(vno_trans, lbltotal, IIf(txtkembali < 0, txtkembali, 0))
        GoTo lanjut
    End If
    
    Exit Sub

lanjut:
    If Linked Then Call Upload_to_Server(vno_trans)
    If txtcard_no <> "CM000-00000" Then Call Save_Point(vno_trans, txtcard_no)
    Call OpenLaci(0) ' buka drawer tanpa print
    Call CetakStruk(vstatus, vno_trans)
    
    If vstatus = "SALES" Then Call CetakPromo(vno_trans)
    
    For x = 0 To 4
        frmSales.cmdsales(x).Enabled = False
    Next x
    frmSales.cmdsales(8).Enabled = False
    frmSales.cmdsales(9).Enabled = False
    
    frmSales.Label1 = "Change : Rp"
    
    If vstatus = "SALES" Then
        frmSales.lblgrand_total = Format(IIf(txtkembali > 0, txtkembali, 0), "#,##0")
    Else
        frmSales.lblgrand_total = Format(IIf(txtkembali < 0, txtkembali, 0), "#,##0")
    End If
    
    frmSales.cmdsales(6).Caption = "VALIDATE TOTAL"
    frmSales.AdoLocal.Recordset.MoveLast
    Call CDisplay("CHANGE :", "Rp. " & frmSales.lblgrand_total)
    
    vpay = 0
    Unload Me
    Exit Sub
        
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Cek Lunas " & Err.Description & " " & Err.Number)
End Sub

Private Sub CetakPromo(No_trans As String)
Dim RsPromo As New ADODB.Recordset, RsBayar As New ADODB.Recordset
Dim JmlKupon As Byte, NilaiOK As Long, Msg As String

    StrSQL = "SELECT promo_hdr.promo_id, promo_name, min_purchase, disc, tipe, voucher, lipat, ismsg, isprn " & _
             ", SUM(Sales_Transaction_Details.Net_Price) As Belanja, islimit, qtylimit, qtyout FROM Promo_Hdr " & _
             "INNER JOIN Promo_Dtl ON Promo_Hdr.promo_id = Promo_Dtl.promo_id " & _
             "INNER JOIN Sales_Transaction_Details ON Promo_Dtl.PLU = Sales_Transaction_Details.PLU " & _
             "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
             "AND getdate() Between Start_Date And End_Date and aktif=1" & _
             "GROUP BY promo_hdr.promo_id, promo_name, min_purchase, disc, tipe, voucher, lipat, ismsg, isprn, " & _
             "islimit, qtylimit, qtyout Having (promo_hdr.tipe > 10)"

    
    If Linked Then
        RsPromo.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsPromo.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
    
    If RsPromo.EOF Then
        RsPromo.Close: Set RsPromo = Nothing
        Exit Sub
    End If
        
    While Not RsPromo.EOF
        Select Case RsPromo!tipe
        Case 11 'GWP
            RsBayar.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                         "From Paid where(Payment_Types <> '8') " & _
                         "GROUP BY Transaction_Number " & _
                         "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            JmlKupon = 0
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If RsBayar!bayar >= RsPromo!min_purchase Then
                    NilaiOK = RsBayar!bayar
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= RsPromo!min_purchase Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If RsBayar!bayar < RsPromo!Belanja Then
                    NilaiOK = RsBayar!bayar
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= RsPromo!min_purchase Then
                    JmlKupon = roundDown(NilaiOK / RsPromo!min_purchase)
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsBayar!bayar < RsPromo!Belanja Then
                    NilaiOK = RsBayar!bayar
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                    
                If NilaiOK >= RsPromo!min_purchase Then
                    JmlKupon = roundDown(NilaiOK / RsPromo!min_purchase)
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!qtylimit < RsPromo!qtyout + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut1
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", 00)"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                Msg = "Anda mendapatkan :"
                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_name, No_trans, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            End If
lanjut1:
            RsBayar.Close: Set RsBayar = Nothing
            
        Case 12 'Undian
            RsBayar.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                 "From Paid where(Payment_Types <> '8') " & _
                 "GROUP BY Transaction_Number " & _
                 "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsBayar!bayar >= RsPromo!min_purchase Then
                    NilaiOK = RsBayar!bayar
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= RsPromo!min_purchase Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                Msg = RsPromo!promo_name
                Msg = Msg + vbNewLine + "Nama :"
                Msg = Msg + vbNewLine + vbNewLine + "No HP :"
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Masukan struk ini di Information Counter"
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_name, No_trans, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            End If
lanjut2:
            RsBayar.Close: Set RsBayar = Nothing
            
        End Select
        RsPromo.MoveNext
    Wend
    RsPromo.Close: Set RsPromo = Nothing
End Sub

Private Function Pakai_Voc(No_trans As String) As Long
Dim RsVoc As New ADODB.Recordset

    RsVoc.Open "select isnull(sum(Paid_Amount),0) as Nvoc from paid pd inner join Payment_Types pt on pd.Payment_Types = pt.Payment_Types " & _
               "where types in ('SV','OV') and Transaction_Number = '" & No_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Pakai_Voc = RsVoc!NVoc
    RsVoc.Close: Set RsVoc = Nothing
End Function

Private Sub Paid_To_Sales(nomor As String, total As Long, kembali As Long)
Dim RsHitung As New ADODB.Recordset
On Error GoTo ErrH

    RsHitung.Open "select sum(discount_amount+extradisc_amt) as hemat " & _
                  "from sales_transaction_details where transaction_number='" & nomor & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    StrSQL = "update sales_transactions set total_paid =" & total + kembali & ", change_amount=" & kembali & _
            ", total_discount=" & RsHitung!hemat & ", status='00' , net_price=net_amount where transaction_number = '" & nomor & "'"
    
    ConnLocal.Execute StrSQL
    
    RsHitung.Close: Set RsHitung = Nothing
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Paid_To_Sales " & Err.Description & " " & Err.Number)
End Sub

Private Sub Upload_to_Server(nomor As String)
On Error GoTo ErrH
Dim Dbs As String, Svr As String

    Svr = "[" & VSvr & "]"
    Dbs = Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini")
    
    'ConnLocal.Execute "exec spp_UploadDataToServer '" & nomor & "','" & Svr & "',''"
    StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.Sales_Transactions (Transaction_Number, Cashier_ID, Customer_ID, Card_Number, Spending_Program_ID, " & _
        "Transaction_Date, Total_Discount, Points_Of_Spending_Program, Point_Of_Item_Program, Point_Of_Card_Program, " & _
        "Payment_Program_ID, Branch_ID, Cash_Register_ID, Total_Paid, Net_Price, Tax, Net_Amount, Change_Amount, " & _
        "Flag_Arrange, WorkManShip, Flag_Return, Register_Return, Transaction_Date_Return, Transaction_Number_Return, " & _
        "Last_Point, Get_Point, Status, Upload_Status, Transaction_Time, Store_Type) " & _
        "(SELECT  Transaction_Number, Cashier_ID, Customer_ID, Card_Number, Spending_Program_ID, Transaction_Date, Total_Discount, Points_Of_Spending_Program, " & _
        "Point_Of_Item_Program, Point_Of_Card_Program, Payment_Program_ID, Branch_ID, Cash_Register_ID, Total_Paid, Net_Price, Tax, Net_Amount, Change_Amount, " & _
        "Flag_Arrange, WorkManShip, Flag_Return, Register_Return, Transaction_Date_Return, Transaction_Number_Return, Last_Point, Get_Point, Status, Upload_Status, " & _
        "Transaction_Time , Store_Type " & _
        "FROM Sales_Transactions where Transaction_Number='" & nomor & "')"
    
    ConnLocal.Execute StrSQL
    
    ConnLocal.Execute "Insert " & Svr & "." & Dbs & ".dbo.Sales_Transaction_details (Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
        "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, " & _
        "Flag_Status, Flag_Paket_Discount) " & _
        "(select Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
        "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, " & _
        "Flag_Status , Flag_Paket_Discount " & _
        "FROM [" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & "]." & _
        Cfg_Get("Local", "DatabaseName", App.Path & "\config.ini") & ".dbo.Sales_Transaction_details where Transaction_Number='" & nomor & "')"

    ConnLocal.Execute "Insert " & Svr & "." & Dbs & ".dbo.paid(Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, " & _
        "Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) " & _
        "(select Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No, " & _
        "Credit_Card_Name , Paid_Amount, Shift " & _
        "FROM [" & Cfg_Get("Local", "ServerName", App.Path & "\config.ini") & "]." & _
        Cfg_Get("Local", "DatabaseName", App.Path & "\config.ini") & ".dbo.paid where Transaction_Number='" & nomor & "')"

    ConnLocal.Execute "Update sales_transactions set upload_status='99' where transaction_number='" & nomor & "'"

    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Upload_to_Server " & Err.Description & " " & Err.Number)
End Sub

Private Sub tampil_point()
    txtsaldo_point = Star_Pt
    txtnilai_point = Format(txtsaldo_point * txtharga_point, "#,##0")
End Sub

Private Sub txttukar_point_Change()
    txtpoint = Format(txttukar_point * txtharga_point, "#,##0")
End Sub

Private Sub txtvoucher_InvalidInput()
    MsgBox "Angka tidak valid", vbOKOnly + vbInformation, "Oops.."
    txtvoucher = 0
    txtvoucher.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub txtcash_InvalidInput()
    MsgBox "Angka tidak valid", vbOKOnly + vbInformation, "Oops.."
    txtcash = 0
    txtcash.SetFocus
    SendKeys "{home}+{end}"
End Sub
