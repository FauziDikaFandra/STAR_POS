VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10860
   ControlBox      =   0   'False
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid80.TDBGrid Grid1 
      Bindings        =   "frmSales.frx":08CA
      Height          =   4260
      Left            =   150
      TabIndex        =   37
      Top             =   1500
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   7514
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
      Columns(1).Caption=   "SPG"
      Columns(1).DataField=   "Flag_Paket_Discount"
      Columns(1).DataWidth=   11
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PLU"
      Columns(2).DataField=   "PLU"
      Columns(2).DataWidth=   18
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Deskripsi"
      Columns(3).DataField=   "Item_Description"
      Columns(3).DataWidth=   50
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Harga"
      Columns(4).DataField=   "Price"
      Columns(4).DataWidth=   22
      Columns(4).NumberFormat=   "#,##0"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Qty"
      Columns(5).DataField=   "Qty"
      Columns(5).DataWidth=   23
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Disc1"
      Columns(6).DataField=   "Discount_Percentage"
      Columns(6).DataWidth=   23
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Disc1 Rp"
      Columns(7).DataField=   "Discount_Amount"
      Columns(7).DataWidth=   22
      Columns(7).NumberFormat=   "#,##0"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Disc2"
      Columns(8).DataField=   "ExtraDisc_Pct"
      Columns(8).DataWidth=   23
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Disc2 Rp"
      Columns(9).DataField=   "ExtraDisc_Amt"
      Columns(9).DataWidth=   22
      Columns(9).NumberFormat=   "#,##0"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Jumlah"
      Columns(10).DataField=   "Net_Price"
      Columns(10).DataWidth=   22
      Columns(10).NumberFormat=   "#,##0"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "flag_void"
      Columns(11).DataField=   "flag_void"
      Columns(11).DataWidth=   11
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).AllowFocus=   0   'False
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=423"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1138"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1032"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._AlignLeft=0"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2858"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2752"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=3995"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3889"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1614"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1508"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(4)._AlignLeft=0"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=794"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=688"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(5)._AlignLeft=0"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=1058"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=953"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(6)._AlignLeft=0"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=1323"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1217"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(7)._AlignLeft=0"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=873"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=767"
      Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(43)=   "Column(8)._AlignLeft=0"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=1323"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=1217"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(48)=   "Column(9)._AlignLeft=0"
      Splits(0)._ColumnProps(49)=   "Column(10).Width=1799"
      Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=1693"
      Splits(0)._ColumnProps(52)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(53)=   "Column(10)._AlignLeft=0"
      Splits(0)._ColumnProps(54)=   "Column(11).Width=582"
      Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=476"
      Splits(0)._ColumnProps(57)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(59)=   "Column(11)._AlignLeft=0"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(26)  =   ":id=22,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(27)  =   ":id=22,.fontname=MS Sans Serif"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(86)  =   "Named:id=33:Normal"
      _StyleDefs(87)  =   ":id=33,.parent=0"
      _StyleDefs(88)  =   "Named:id=34:Heading"
      _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=34,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=35:Footing"
      _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=36:Selected"
      _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=37:Caption"
      _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(97)  =   "Named:id=38:HighlightRow"
      _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=39:EvenRow"
      _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=40:OddRow"
      _StyleDefs(102) =   ":id=40,.parent=33"
      _StyleDefs(103) =   "Named:id=41:RecordSelector"
      _StyleDefs(104) =   ":id=41,.parent=34"
      _StyleDefs(105) =   "Named:id=42:FilterBar"
      _StyleDefs(106) =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   150
      TabIndex        =   33
      Top             =   5925
      Width           =   10590
      Begin VB.CommandButton CmdNav 
         Caption         =   "ENTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   0
         Left            =   3375
         Picture         =   "frmSales.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   990
      End
      Begin VB.CommandButton CmdNav 
         Caption         =   "&NUM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Index           =   3
         Left            =   4425
         Picture         =   "frmSales.frx":12E3
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Width           =   990
      End
      Begin TDBText6Ctl.TDBText txtkode 
         Height          =   390
         Left            =   150
         TabIndex        =   0
         Top             =   105
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   688
         Caption         =   "frmSales.frx":186D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSales.frx":18D7
         Key             =   "frmSales.frx":18F5
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
         Format          =   "9A@"
         FormatMode      =   0
         AutoConvert     =   0
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
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "TOTAL : Rp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   5475
         TabIndex        =   35
         Top             =   75
         Width           =   2415
      End
      Begin VB.Label lblgrand_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   7050
         TabIndex        =   34
         Top             =   0
         Width           =   3465
      End
   End
   Begin MSAdodcLib.Adodc AdoLocal 
      Height          =   330
      Left            =   8625
      Top             =   3825
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUSTOMER "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   150
      TabIndex        =   10
      Top             =   75
      Width           =   6990
      Begin VB.CommandButton cmdsales 
         Caption         =   "MEMBER CARD"
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
         Index           =   8
         Left            =   5820
         Picture         =   "frmSales.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   300
         Width           =   1020
      End
      Begin TDBText6Ctl.TDBText txtcard_no 
         Height          =   390
         Left            =   150
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   688
         Caption         =   "frmSales.frx":232B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSales.frx":23A7
         Key             =   "frmSales.frx":23C5
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
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
         Text            =   "CM000-00000"
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
      Begin TDBText6Ctl.TDBText txtcust_name 
         Height          =   390
         Left            =   150
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   750
         Width           =   5565
         _Version        =   65536
         _ExtentX        =   9816
         _ExtentY        =   688
         Caption         =   "frmSales.frx":2417
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSales.frx":2491
         Key             =   "frmSales.frx":24AF
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
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   "ONE TIME CUSTOMER"
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
      Begin TDBText6Ctl.TDBText txtpoint 
         Height          =   390
         Left            =   4275
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   688
         Caption         =   "frmSales.frx":2501
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSales.frx":256B
         Key             =   "frmSales.frx":2589
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
         MaxLength       =   5
         LengthAsByte    =   0
         Text            =   "0"
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
      Begin TDBText6Ctl.TDBText txtcust_id 
         Height          =   390
         Left            =   7050
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   300
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   688
         Caption         =   "frmSales.frx":25DB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmSales.frx":263B
         Key             =   "frmSales.frx":2659
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
         MaxLength       =   30
         LengthAsByte    =   0
         Text            =   "100000"
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
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   150
      TabIndex        =   1
      Top             =   6675
      Width           =   10590
      Begin VB.CommandButton CmdNav 
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
         Left            =   8475
         Picture         =   "frmSales.frx":26AB
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton CmdNav 
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
         Left            =   7425
         Picture         =   "frmSales.frx":30AD
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
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
         Index           =   7
         Left            =   9525
         Picture         =   "frmSales.frx":3AAF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "ESC"
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
         Index           =   6
         Left            =   6375
         Picture         =   "frmSales.frx":44B1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "VALIDATE ARTICLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   5
         Left            =   5325
         Picture         =   "frmSales.frx":4EB3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "VIEW"
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
         Index           =   4
         Left            =   4275
         Picture         =   "frmSales.frx":58B5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "LINE VOID"
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
         Index           =   3
         Left            =   3225
         Picture         =   "frmSales.frx":62B7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "DISCOUNT"
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
         Left            =   2175
         Picture         =   "frmSales.frx":6CB9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "HOLD"
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
         Left            =   1125
         Picture         =   "frmSales.frx":76BB
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdsales 
         Caption         =   "PAYMENT"
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
         Left            =   75
         Picture         =   "frmSales.frx":80BD
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdsales 
      Caption         =   "CEK PROMO"
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
      Index           =   9
      Left            =   7200
      Picture         =   "frmSales.frx":8ABF
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   375
      Width           =   1020
   End
   Begin VB.Label v_burui 
      Caption         =   "Label7"
      Height          =   135
      Left            =   9480
      TabIndex        =   46
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total QTY :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8475
      TabIndex        =   42
      Top             =   1050
      Width           =   1290
   End
   Begin VB.Label lblqty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   9675
      TabIndex        =   41
      Top             =   1050
      Width           =   540
   End
   Begin VB.Label vpromo 
      Caption         =   "promo"
      Height          =   315
      Left            =   5400
      TabIndex        =   39
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label lblno 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   8475
      TabIndex        =   38
      Top             =   75
      Width           =   1965
   End
   Begin VB.Label vgtotal 
      Caption         =   "gtotal"
      Height          =   315
      Left            =   6450
      TabIndex        =   31
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label vtotal 
      Caption         =   "total"
      Height          =   315
      Left            =   3900
      TabIndex        =   29
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10200
      TabIndex        =   28
      Top             =   750
      Width           =   315
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10200
      TabIndex        =   27
      Top             =   450
      Width           =   315
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Disc :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8475
      TabIndex        =   26
      Top             =   750
      Width           =   1290
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8475
      TabIndex        =   25
      Top             =   450
      Width           =   1290
   End
   Begin VB.Label vdiscrp2 
      Caption         =   "Disc2Rp"
      Height          =   315
      Left            =   3150
      TabIndex        =   24
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label vdiscrp1 
      Caption         =   "Disc1Rp"
      Height          =   315
      Left            =   2400
      TabIndex        =   23
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label vdisc2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   9675
      TabIndex        =   22
      Top             =   750
      Width           =   540
   End
   Begin VB.Label vqty 
      Caption         =   "qty"
      Height          =   315
      Left            =   225
      TabIndex        =   21
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label vdisc1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   9675
      TabIndex        =   20
      Top             =   450
      Width           =   540
   End
   Begin VB.Label vflag 
      Caption         =   "flag"
      Height          =   315
      Left            =   4650
      TabIndex        =   19
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label vno_trans 
      Height          =   315
      Left            =   225
      TabIndex        =   18
      Top             =   3825
      Width           =   1890
   End
   Begin VB.Label vharga 
      Caption         =   "harga"
      Height          =   315
      Left            =   975
      TabIndex        =   17
      Top             =   4200
      Width           =   1365
   End
   Begin VB.Label vdesc 
      Caption         =   "desc"
      Height          =   315
      Left            =   4425
      TabIndex        =   16
      Top             =   3825
      Width           =   2715
   End
   Begin VB.Label vspg 
      Caption         =   "spg"
      Height          =   315
      Left            =   2175
      TabIndex        =   15
      Top             =   3825
      Width           =   690
   End
   Begin VB.Label vplu 
      Caption         =   "plu"
      Height          =   315
      Left            =   2925
      TabIndex        =   14
      Top             =   3825
      Width           =   1440
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Qty() As String

Dim PLU_SPG As String
Dim VTukar_Point As String
Dim DiscNonKaryawan As Byte
Dim TxtNom As Integer
Dim xdisc_value As Long
Dim Promo4Ever As Boolean
Dim OnlyOncePromoTipe20 As Boolean
Dim InfoPromox, HasilPromox, PromoIDx, CekPromoIDx As String
Dim LimitTipe9 As Long

Private Sub Update_MySTAR()
    Dim RsMem As New ADODB.Recordset
            
    RsMem.Open "select card_number, point_of_card_program from sales_transactions where transaction_number = '" & _
                vno_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    If Not RsMem.EOF Then Call MySTAR(Trim(RsMem!card_number), 0)
    txtcard_no = UCase(RsMem!card_number)
    Star_No = UCase(RsMem!card_number)
    
    txtcust_name = Star_Nm
    txtcust_id = Star_Id
    txtpoint = Star_Pt
    VBonus_Point = RsMem!Point_Of_Card_Program
    RsMem.Close: Set RsMem = Nothing
End Sub

Private Sub CmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'Enter
            SendKeys "{Enter}"
        Case 1 'Up
            Grid1.MovePrevious
        Case 2 'Down
            Grid1.MoveNext
        Case 3 ' Num
            frmNum.Caption = "NUMBER - SALES"
            frmNum.Show 1
    End Select
    txtkode.SetFocus
End Sub

Private Sub Form_Activate()
    Frame3.Caption = "CUSTOMER - " & VBonus_Point
End Sub

Private Sub Form_Load()
    Call Kosong
    Promo4Ever = False
    UpdateStatusSeqDetail = False
    SeqCountInt = 0
    cmdsales(9).Enabled = VAda_Promo
    vno_trans = VNomor
    PLU_SPG = ""
    VKary = ""
    VTanya = False
    OnlyCheckPromo = 0
    OnlyOncePromoTipe20 = False
    If vno_trans <> "" Then
        Update_MySTAR
    Else
        vno_trans = Gen_No
    End If
    isLimitCC = 0
    AdoLocal.ConnectionString = StrConLoc
    AdoLocal.RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                            "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void FROM Sales_Transaction_Details " & _
                            "where transaction_number='" & vno_trans & "'"
    AdoLocal.Refresh
        
    vgtotal = 0
    lblqty = 0

    While Not AdoLocal.Recordset.EOF
        vgtotal = vgtotal + AdoLocal.Recordset!Net_Price
        lblqty = Val(lblqty) + AdoLocal.Recordset!Qty
        AdoLocal.Recordset.MoveNext
    Wend
    
    lblgrand_total = Format(vgtotal, "#,##0")
End Sub

Private Sub cmdsales_Click(Index As Integer)
    Select Case Index
    Case 0 'Payment
        If AdoLocal.Recordset.RecordCount = 0 Then Exit Sub
        OnlyCheckPromo = 1
        If cmdsales(9).Enabled Then cmdsales_Click (9)
        'dihilangkan dulu agak berat
        'If Star_No <> "CM000-00000" And Me.Caption = "SALES" Then Call Bayar_Point
        '-------
        'MsgBox VNomor
        VNomor = vno_trans
        If StrukEmail = True And Me.Caption = "SALES" Then
            frmDataCustomer.vtotalx = vgtotal
            frmDataCustomer.vdiscx = VDiscBySTAR
            frmDataCustomer.txtcardx = txtcard_no
            frmDataCustomer.txtcapt = Me.Caption
            frmDataCustomer.Show
        Else
            With frmPayment
            Call CDisplay("TOTAL :", "Rp. " & Format(vgtotal - VDiscBySTAR, "#,##0"))
            .vpay = vgtotal - VDiscBySTAR
            .txtcard_no = txtcard_no
            .vstatus = Me.Caption
            .Show 1
        End With
        txtkode = ""
        txtkode.SetFocus
        End If
        
    Case 1 'Hold
        If AdoLocal.Recordset.RecordCount = 0 Then Exit Sub
        Call Simpan_Header
        Call CetakPesan("HOLD", vno_trans)
        Call SaveLog("Hold Transaction " & vno_trans & " " & VKasir_ID & " / " & VKasir_Nm)
        VNomor = ""
        Unload Me
        frmMain.Show
    Case 2 'Discount
        If vdisc2 = 0 Then
            frmDisc.lblmsg.Caption = "DISCOUNT"
            frmDisc.Show 1
        Else
            MsgBox "Key Discount Maksimum 2 X", vbOKOnly + vbInformation, "Oops.."
        End If
        txtkode.SetFocus
    Case 3 'Line Void
        If Grid1.Columns(0) <> "" Then
            'flag void <> 1 dan qty> 0
            If Grid1.Columns(11) = 1 Or Grid1.Columns(5) < 0 Then
                MsgBox "Item tersebut tidak dapat divoid", vbOKOnly + vbInformation, "Oops.."
                Exit Sub
            End If
            
            If Not Super(2) Then Exit Sub
            
            With Grid1
                ConnLocal.Execute "INSERT INTO Sales_Transaction_Details(Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
                "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, Flag_Status, Flag_Paket_Discount) " & _
                "VALUES ('" & vno_trans & "','" & Gen_Seq & "','" & .Columns(2) & "','" & .Columns(3) & "','" & .Columns(4) & "','" & -1 * .Columns(5) & _
                "','" & -1 * .Columns(4) * .Columns(5) & "','" & .Columns(6) & "','" & -1 * .Columns(7) & "','" & .Columns(8) & "','" & -1 * .Columns(9) & "','" & -1 * .Columns(10) & _
                "','0','0','0','" & .Columns(1) & "')"
                            
                ConnLocal.Execute "update Sales_Transaction_Details set flag_void=1 where transaction_number='" & vno_trans & "' and seq='" & .Columns(0) & "'"

                vgtotal = vgtotal - Val(.Columns(10).Value)
                lblqty = Val(lblqty) - Val(.Columns(5).Value)
                lblgrand_total = Format(vgtotal, "#,##0")
                
                Call Simpan_Header
            End With
            AdoLocal.Refresh
            txtkode.SetFocus
        End If
    Case 4 'View
        If txtkode.Caption = "ID SPG" Then
            MsgBox "Isi dahulu ID SPG", vbOKOnly + vbInformation, "Oops.."
            txtkode.SetFocus
            Exit Sub
        End If
        frmView.Show 1
        DoEvents
        txtkode.SetFocus
        If txtkode <> "" Then SendKeys "{Enter}"
    Case 5 'Validate detail
        If cmdsales(6).Caption = "ESC" Then Exit Sub
        Dim RsVali As New ADODB.Recordset
        Dim LongDesc As String
        With Grid1
            Dim aa As String
            Dim bb As String
            RsVali.Open "select Long_Description from item_master where plu = '" & .Columns(2) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            LongDesc = .Columns(3)
            If Not RsVali.EOF Then LongDesc = RsVali!Long_Description
                aa = .Columns(2) & " " & .Columns(5) & " X " & Format(.Columns(4), "#,##0") & " = Rp." & Format(.Columns(10), "#,##0")
                If .Columns(6) > 0 Then aa = aa & vbNewLine & "Disc." & .Columns(6) & "% = " & Format(.Columns(7), "#,##0")
                If .Columns(8) > 0 Then aa = aa & vbNewLine & "Disc." & .Columns(8) & "% = " & Format(.Columns(9), "#,##0")
                bb = LongDesc & vbNewLine & .Columns(1) & "/" & Siapa_SPG(.Columns(1))
                If Trim(aa) <> "X" Then Call CetakValid(vno_trans, aa, bb)
                Call SaveLog("Validasi Kasir : " & VKasir_ID & "/" & VKasir_Nm & "SPG : " & .Columns(1) & "/" & Siapa_SPG(.Columns(1)))
        End With
        txtkode.SetFocus
    Case 6 'Esc atau validate total
        If cmdsales(6).Caption = "ESC" Then
            Call Kosong
        Else
            'Call CetakValid(vno_trans, "Total Rp. " & Format(vgtotal, "#,##0"), "")
            frmDisc.lblmsg.Caption = "VALIDASI"
            frmDisc.Show 1
        End If
        txtkode.SetFocus
    Case 7 'Close
        VNomor = ""
        If AdoLocal.Recordset.RecordCount > 0 And Label1.Caption = "TOTAL : Rp" Then Call CetakPesan("HOLD", vno_trans)
        Unload Me
        frmMain.Show
    Case 8 'Member
        frmCard.Caption = Me.Caption
        frmCard.Show 1
        DoEvents
        ConnLocal.Execute "UPDATE Sales_Transactions set customer_id = '" & txtcust_id & "', card_number = '" & txtcard_no & _
                "' where Transaction_Number='" & vno_trans & "'"
        txtkode.SetFocus
    Case 9 'Cek Promo
        Dim RsInfoPromo As New ADODB.Recordset
        Dim RsInfoPromo16 As New ADODB.Recordset
        Dim RsInfoPromoSeq As New ADODB.Recordset
        Dim TotKelipatan As Integer
        'Dim pcsx As String
        InfoPromox = 0
        HasilPromox = 0
        PromoIDx = 0
        CekPromoIDx = 0
        TotKelipatan = 0
        If AdoLocal.Recordset.RecordCount = 0 Then Exit Sub
        If OnlyCheckPromo <> 1 Then
            OnlyCheckPromo = 2
            RsInfoPromo.Open "select txt1 as tipe,txt2 as Kodex,txt3 as kata1,txt4 as kata2 from promo_hdr where tipe='28' and getdate() Between Start_Date And End_Date and aktif=1", ConnLocal, adOpenForwardOnly, adLockReadOnly
            If Not RsInfoPromo.EOF Then
                InfoPromox = RsInfoPromo!Tipe
                PromoIDx = RsInfoPromo!Kodex
            End If
        End If
        
        Call Cek_Promo
        If VPing = "ONLINE" Then
            If Star_No <> "CM000-00000" Then
                SeqCountInt = 0
                RsInfoPromoSeq.Open "select a.promo_id from Promo_Hdr a inner join Seqmentation_Member_Promo b on a.promo_id = b.promo_id where b.card_nr " & _
                " = '" & Star_No & "' and getdate() Between a.Start_Date And a.End_Date and a.aktif=1 and b.status = 1 and a.tipe < 31 and a.seqmentation <> 0", ConnServer, adOpenForwardOnly, adLockReadOnly
                While Not RsInfoPromoSeq.EOF
                    UpdateStatusSeqDetail = True
                    Call Cek_Promo_Seq(RsInfoPromoSeq!promo_id)
                    
                    RsInfoPromoSeq.MoveNext
                Wend
                RsInfoPromoSeq.Close: Set RsInfoPromoSeq = Nothing
            End If
        End If
        
        
        If HasilPromox = 1 Then
        MsgBox RsInfoPromo!kata1 & RsInfoPromo!kata2, vbOKOnly + vbInformation, "Oops.."
        End If
        
        RsInfoPromo16.Open "select c.promo_id,c.promo_name, c.min_purchase, c.txt1,c.txt2,c.txt3,c.txt4,c.lipat, " & _
                            "SUM(a.net_price) As Total from Sales_Transaction_Details a inner join Promo_Dtl b on a.PLU = b.PLU " & _
                            "inner join Promo_Hdr c on b.promo_id = c.promo_id where getdate() Between c.Start_Date And c.End_Date " & _
                            "and c.aktif=1 and a.Transaction_Number = '" & vno_trans & "' and c.tipe = '16' group by c.promo_id," & _
                            "c.promo_name, c.min_purchase, c.min_member, c.txt1,c.txt2,c.txt3,c.txt4,c.lipat order by c.promo_id", ConnLocal, adOpenForwardOnly, adLockReadOnly
        
        While Not RsInfoPromo16.EOF
            If RsInfoPromo16!min_purchase <= RsInfoPromo16!total Then
                TotKelipatan = roundDown(RsInfoPromo16!total / RsInfoPromo16!min_purchase)
                MsgBox RsInfoPromo16!txt1 & " " & TotKelipatan & " Pcs " & RsInfoPromo16!txt2 & " " & RsInfoPromo16!txt3, vbOKOnly + vbInformation, "Oops.."
            End If
            RsInfoPromo16.MoveNext
        Wend
        RsInfoPromo16.Close: Set RsInfoPromo16 = Nothing
        
        Call Simpan_Header
        Call CDisplay("TOTAL :", "Rp. " & Format(vgtotal - VDiscBySTAR, "#,##0"))
    End Select
End Sub

Private Sub Bayar_Point()
If Promo4Ever = True Then
Promo4Ever = False
Exit Sub
End If
Dim RsByrPoint As New ADODB.Recordset
    
    RsByrPoint.Open "select SUM(qty*points_received) as byr_pt, SUM(qty*flag_void) as byr_rp from Sales_Transaction_Details where transaction_number='" & _
                    vno_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                    
    If Not RsByrPoint.EOF And RsByrPoint!byr_pt > 2 Then
        If RsByrPoint!byr_pt < Star_Pt Then
                VTukar_Point = Pay_Point(RsByrPoint!byr_pt, Star_No, vno_trans, RsByrPoint!byr_rp)
                If VTukar_Point <> "GAGAL" Then
                    ConnLocal.Execute "Delete from paid where transaction_number='" & vno_trans & "'"
                
                    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
                        "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
                        vno_trans & "','5','1','IDR','1','" & Star_No & "','" & VTukar_Point & "'," & RsByrPoint!byr_rp & ",'" & VShift & "')"
                
                VDiscBySTAR = RsByrPoint!byr_rp
            Else
                MsgBox "Pembayaran dengan point reward GAGAL", vbOKOnly + vbInformation, "Oops.."
            End If
        Else
            MsgBox "Point reward tidak mencukupi untuk pembayaran", vbOKOnly + vbInformation, "Oops.."
        End If
    End If
    RsByrPoint.Close: Set RsByrPoint = Nothing
End Sub

Private Function Cari_Item_Rewards(kode) As String
Dim RsCari As New ADODB.Recordset
            
    RsCari.Open "select plu, point, rupiah from item_rewards where plu = '" & _
                Trim(kode) & "' and getdate() Between Start_Date And End_Date and aktif=1 ", ConnServer, adOpenForwardOnly, adLockReadOnly

    If Not RsCari.EOF Then
        If MsgBox("Apakah item ini akan dibayar dengan point rewards?", vbYesNo + vbOKOnly, "Oops..") = vbYes Then
            
            ConnLocal.Execute "update sales_transaction_details set points_received='" & RsCari!Point & _
                            "' , flag_void='" & RsCari!rupiah & "' where plu='" & kode & _
                            "' and transaction_number='" & vno_trans & "'"
        End If
    End If

    RsCari.Close: Set RsCari = Nothing
End Function

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Or KeyAscii = 39 Then KeyAscii = 0    '45 - dan '39 '
End Sub


Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Cmd As String
Dim RsX As New ADODB.Recordset

    Select Case KeyCode
    Case 13
        Select Case txtkode.Caption
        Case "ID SPG"
            If frmSales.cmdsales(6).Caption <> "ESC" Then Exit Sub
            
            If txtkode.Text = "0" Then
                If Cari_SPG("0") = True Then
                    txtkode.Caption = "PLU"
                    txtkode = ""
                    Exit Sub
                End If
            End If
            'If Left(txtkode.Text, 2) = "24" And Len(txtkode.Text) = 13 Then
            If Len(txtkode.Text) = 9 Then
                'If Cari_SPG(CDbl(Mid(txtkode.Text, 3, 10))) = True Then
                If Cari_SPG(CDbl(txtkode.Text)) = True Then
                    PLU_SPG = txtkode
                    txtkode.Caption = "PLU"
                    txtkode = ""
                End If
            ElseIf Len(txtkode.Text) <= 5 And Len(txtkode.Text) > 2 Then
                If Cari_SPG(txtkode.Text) = True Then
                    PLU_SPG = txtkode
                    txtkode.Caption = "PLU"
                    txtkode = ""
                End If
            Else
                MsgBox "ID SPG tidak terdaftar", vbOKOnly + vbInformation, "Oops.."
            End If

        Case "PLU"
            Dim CekPLU() As String
            Dim RealPLU As String
            CekPLU = Split(txtkode.Text, "*")
            On Error GoTo ErrH
            If Len(CekPLU(0)) > 10 Then
                RealPLU = CekPLU(0)
            Else
                RealPLU = CekPLU(1)
            End If
            If Cari_PLU(Right(Trim(RealPLU), 14)) = True Then
            
                Qty = Split(txtkode.Text, "*")
                vqty = IIf(Len(Qty(0)) > 10, 1, Qty(0))
                If vqty = "" Then vqty = 1
                If Me.Caption = "REFUND" Then vqty = vqty * -1
                
                'vflag = 0 Barang Direct, vflag = 1 Barang Konsinyasi
                If v_burui = 0 And vdisc1 <> 0 And vpromo <> "disc" Then
                    If Not Super(1) Then Exit Sub
                End If
    
                VOK = False
                If vflag = 1 Then frmHarga.Show 1
                DoEvents
                
                vdiscrp1 = Format(vqty * vharga * vdisc1 / 100, "#,##0")
                vdiscrp2 = Format((vqty * vharga - vdiscrp1) * vdisc2 / 100, "#,##0")
                vtotal = vqty * vharga - vdiscrp1 - vdiscrp2
                
                If VOK = True Or vflag = 0 Then
                    vgtotal = Val(vgtotal) + Val(vtotal)
                    lblqty = Val(lblqty) + vqty
                    lblgrand_total = Format(vgtotal, "#,##0")
                    RsX.Open "SELECT * from paid where transaction_number = '" & vno_trans & "'", ConnLocal, adOpenStatic, adLockReadOnly
                    If Not RsX.EOF Then
                        RsX.Close: Set RsX = Nothing
                        DoEvents
                        SendKeys "{home}+{end}"
                        txtkode.SetFocus
                        Exit Sub
                    End If
                    RsX.Close: Set RsX = Nothing
                    Call Simpan_Header
                    Call Simpan_Detail
                    Call CDisplay(Left(vdesc, 20), Left(CStr(vqty) & " pcs Rp. " & Format(vtotal, "#,##0"), 20))
                    'PEMBAYARAN DENGAN POINT REWARDS
                    If Star_No <> "CM000-00000" Then
                        If Linked Then Call Cari_Item_Rewards(txtkode.Text)
                    End If
                End If
                
                AdoLocal.RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void FROM Sales_Transaction_Details " & _
                                        "where transaction_number='" & vno_trans & "'"
                AdoLocal.Refresh
                If AdoLocal.Recordset.RecordCount > 0 Then Grid1.Bookmark = AdoLocal.Recordset.RecordCount
                Call Kosong
                txtkode.Text = PLU_SPG
            End If
        End Select
        DoEvents
        SendKeys "{home}+{end}"
        txtkode.SetFocus
        Exit Sub
ErrH:
        MsgBox "PLU tidak terdaftar", vbOKOnly + vbInformation, "Oops.."
        DoEvents
        SendKeys "{home}+{end}"
        txtkode.SetFocus
    Case 27
        Call Kosong
    Case 38
        Grid1.MovePrevious
    Case 40
        Grid1.MoveNext
    Case Else
    On Error Resume Next
        Cmd = KeyStroke(KeyCode)
    End Select
        
    Select Case Cmd
    Case "DISCOUNT"
        If cmdsales(2).Enabled Then cmdsales_Click (2)
    Case "END"
        If cmdsales(7).Enabled Then cmdsales_Click (7)
    Case "ESC", "VALIDTOT"
        If cmdsales(6).Enabled Then cmdsales_Click (6)
    Case "HOLD"
        If cmdsales(1).Enabled Then cmdsales_Click (1)
    Case "LINE_VOID"
        If cmdsales(3).Enabled Then cmdsales_Click (3)
    Case "PAYMENT"
        If cmdsales(0).Enabled Then cmdsales_Click (0)
    Case "VALIDATE"
        If cmdsales(5).Enabled Then cmdsales_Click (5)
    Case "VIEW"
        If cmdsales(4).Enabled Then cmdsales_Click (4)
    Case "MEMBER"
        If cmdsales(8).Enabled Then cmdsales_Click (8)
    Case "PROMO"
        If cmdsales(9).Enabled Then cmdsales_Click (9)
    End Select
End Sub

Private Function Cari_SPG(kode) As Boolean
Dim RsCari As New ADODB.Recordset
        
    Cari_SPG = False
            
    'RsCari.Open "select spg_id, spg_name from spg where spg_id = '" & _
                'Trim(kode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                
    RsCari.Open "select spg_id, spg_name from spg where spg_barcode = '" & _
                Trim(kode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

    If Not RsCari.EOF Then
        vspg = RsCari!spg_id
        Cari_SPG = True
    Else
        vspg = ""
        MsgBox "ID SPG tidak terdaftar", vbOKOnly + vbInformation, "Oops.."
    End If
    RsCari.Close: Set RsCari = Nothing
End Function

Private Function Siapa_SPG(kode) As String
Dim RsCari As New ADODB.Recordset
            
    'RsCari.Open "select spg_id, spg_name from spg where spg_id = '" & _
                'Trim(kode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                
    RsCari.Open "select spg_id, spg_name from spg where spg_barcode = '" & _
                Trim(kode) & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

    Siapa_SPG = IIf(Not RsCari.EOF, RsCari!spg_name, "")

    RsCari.Close: Set RsCari = Nothing
End Function

Private Function Cari_PLU(kode) As Boolean
Dim RsCari As New ADODB.Recordset

    Cari_PLU = False
    vpromo = ""
    StrSQL = "select Plu,Description,Normal_Price,Current_Price,Flag, disc_percent,burui from item_master where plu = '" & kode & "' and description <> 'TIDAK AKTIF'"
            
    If Linked Then
        RsCari.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsCari.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not RsCari.EOF Then
        vplu = RsCari!plu
        vdesc = RsCari!Description
        vharga = IIf(RsCari!flag = 0, RsCari!current_Price, 0)
        vflag = RsCari!flag
        v_burui = Right(RsCari!burui, 1)
        If RsCari!disc_percent > 0 Then
            vpromo = "disc"
            vdisc1 = RsCari!disc_percent
        End If
        Cari_PLU = True
    Else
        vplu = ""
        vdesc = ""
        vharga = 0
        vflag = ""
        MsgBox "PLU tidak terdaftar", vbOKOnly + vbInformation, "Oops.."
    End If
    RsCari.Close: Set RsCari = Nothing
End Function

Private Sub Simpan_Detail()
On Error GoTo ErrH
    ConnLocal.Execute "INSERT INTO Sales_Transaction_Details(Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
             "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, Flag_Status, Flag_Paket_Discount) " & _
             "VALUES ('" & vno_trans & "','" & Gen_Seq & "','" & vplu & "','" & UbahChar(vdesc) & "','" & vharga & "','" & vqty & _
             "','" & vqty * vharga & "','" & vdisc1 & "','" & vdiscrp1 & "','" & vdisc2 & "','" & vdiscrp2 & "','" & vtotal & _
             "','0','0','0','" & vspg & "')"
    
    Exit Sub

ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Simpan_Detail " & Err.Description & " " & Err.Number)
End Sub

Private Sub Simpan_Header()
On Error GoTo ErrH
Dim RsCari As New ADODB.Recordset
Dim sttStrukEmail As Integer
sttStrukEmail = 0
If StrukEmail = True Then
    sttStrukEmail = 1
End If

    
    RsCari.Open "SELECT Transaction_Number FROM Sales_Transactions where transaction_number='" & _
                vno_trans & "' ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If RsCari.EOF Then
        StrSQL = "INSERT INTO Sales_Transactions(Transaction_Number, Cashier_ID, Customer_ID, Card_Number, Spending_Program_ID, Transaction_Date, Total_Discount, " & _
                "Points_Of_Spending_Program, Point_Of_Item_Program, Point_Of_Card_Program, Payment_Program_ID, Branch_ID, Cash_Register_ID, Total_Paid, Net_Price, Tax, " & _
                "Net_Amount, Change_Amount, Flag_Arrange, WorkManShip, Flag_Return, Register_Return, Transaction_Date_Return, Transaction_Number_Return, Last_Point, " & _
                "Get_Point, Status, Upload_Status, Transaction_Time, Store_Type)" & _
                "VALUES ('" & vno_trans & "','" & VKasir_ID & "','" & txtcust_id & "','" & txtcard_no & "','0','" & Format(GetSrvDate, "YYYY-MM-DD") & _
                "',0,0,0," & VBonus_Point & ",'" & sttStrukEmail & "','" & VBranch_ID & "','" & VReg_ID & "',0,0,0," & vgtotal & ",0,0,0,'" & _
                IIf(Me.Caption = "SALES", 0, 1) & "','',NULL,'',0,0,'01','00','" & Format(GetSrvDate, "HH:NN") & "','1')"
    Else
        StrSQL = "update sales_transactions set net_amount=" & vgtotal & ", card_number='" & txtcard_no & _
                 "', customer_id='" & txtcust_id & "', Point_Of_Card_Program=" & VBonus_Point & " where transaction_number='" & vno_trans & "' "
    End If
0
    ConnLocal.Execute StrSQL
    RsCari.Close:   Set RsCari = Nothing
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Simpan_Header " & Err.Description & " " & Err.Number)
End Sub

Private Function Gen_No() As String
Dim RsCari As New ADODB.Recordset

    RsCari.Open "SELECT  max (CAST(RIGHT(Transaction_Number, 4) AS int)) AS nomor " & _
            "FROM Sales_Transactions where LEFT(transaction_number,16)='" & _
            VBranch_ID + VReg_ID + "-" + Format(GetSrvDate, "DDMMYYYY") & "' ", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If IsNull(RsCari!nomor) Then
        Gen_No = VBranch_ID + VReg_ID + "-" + Format(GetSrvDate, "DDMMYYYY") + "-0001"
    Else
        Gen_No = VBranch_ID + VReg_ID + "-" + Format(GetSrvDate, "DDMMYYYY") + "-" + Right("000" + CStr(RsCari!nomor + 1), 4)
    End If
    RsCari.Close:   Set RsCari = Nothing
End Function

Private Function Gen_Seq() As String
Dim RsCari As New ADODB.Recordset
    
    RsCari.Open "select MAX(seq) as urut from Sales_Transaction_Details where Transaction_Number = '" & _
                vno_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Gen_Seq = IIf(Not IsNull(RsCari!urut), RsCari!urut + 1, 1)
    RsCari.Close:   Set RsCari = Nothing
End Function

Private Function Kosong()
    txtkode.Caption = "ID SPG"
    vspg = ""
    vplu = ""
    vdesc = ""
    vqty = 0
    vharga = 0
    vdisc1 = "0"
    vdiscrp1 = "0"
    vdisc2 = "0"
    vdiscrp2 = "0"
    vtotal = 0
    vflag = ""
    txtkode = ""
End Function

Private Sub vno_trans_Change()
    lblno = "TRANS# " & Right(vno_trans, 4)
End Sub

Private Sub Cek_Promo()
Dim RsPromo As New ADODB.Recordset, RsAmbil As New ADODB.Recordset, RsKaryawan As New ADODB.Recordset
Dim xqty As Integer, xharga As Long, xdisc1 As Byte, xdisc1_amt As Long
Dim Pro_tipe As Byte, Pro_Nm As String, Pro_Disc As Byte, Nama_Promo As String
Dim DiscMOP As Byte, Sisa_Bonus As Long, Bonus As Long

        ConnLocal.Execute "update Sales_Transaction_Details set Discount_Amount = 0,Discount_Percentage = 0,ExtraDisc_Amt = 0,ExtraDisc_Pct = 0 where transaction_number='" & vno_trans & "' and Points_Received in (1,2)"
        ConnLocal.Execute "Update Sales_Transaction_Details set Points_Received = 0 where transaction_number='" & vno_trans & "' and Points_Received  in (1,2)"
        With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void  FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        
        vgtotal = 0
        VDiscBySTAR = 0
        DiscMOP = 0
        VCekKartu = ""
        Bonus = 0
        LimitTipe9 = 0
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe not in (3,5,9,15) and Seqmentation = 0")
            
            Select Case Pro_tipe
                Case 0 'tidak ada promo
                    xdisc1 = .Recordset!Discount_Percentage
                Case 2, 11, 14, 17, 18, 19, 26, 28
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                Case 22, 23, 25
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                Case 21
                If txtcard_no <> "CM000-00000" Then
                    .Recordset!ExtraDisc_pct = Pro_Disc
                    .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (Pro_Disc / 100)
                Else
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                End If
'                Case 3 Sebelum 6/6/2016 utuk disc karyawan
'                    xdisc1 = .Recordset!Discount_Percentage
'                    VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
'                    Nama_Promo = Pro_Nm
                Case 4
                    xdisc1 = Pro_Disc
                    If txtcard_no <> "CM000-00000" Then
                        .Recordset!ExtraDisc_pct = IIf(Not IsNumeric(TxtNom), 0, TxtNom)
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * Pro_Disc / 100)) * IIf(Not IsNumeric(TxtNom), 0, TxtNom) / 100
                    Else
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                
                    
                 Case 29
                    xdisc1 = .Recordset!Discount_Percentage
                    If txtcard_no <> "CM000-00000" Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (Pro_Disc / 100)
                    Else
                        .Recordset!ExtraDisc_pct = IIf(Not IsNumeric(TxtNom), 0, TxtNom)
                        .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (IIf(Not IsNumeric(TxtNom), 0, TxtNom) / 100)
                    End If
                'Case 5
                '    xdisc1 = .Recordset!Discount_Percentage
                '    If DiscMOP = 0 Then
                '        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                '            DiscMOP = 1 'yes
                '        Else
                '            DiscMOP = 2 'no
                '            VCekKartu = ""
                '        End If
                '    End If
                    
                '    If DiscMOP = 1 Then
                '        .Recordset!ExtraDisc_pct = Pro_Disc
                '        .Recordset!ExtraDisc_amt = (xqty * xharga - .Recordset!Discount_Amount) * Pro_Disc / 100
                '    ElseIf DiscMOP = 2 Then
                '        .Recordset!ExtraDisc_pct = 0
                '        .Recordset!ExtraDisc_amt = 0
                '        VCekKartu = ""
                '        VDiscBySTAR = 0
                '    End If
                Case 6
                    xdisc1 = Pro_Disc
                Case 7 ' disc 10% untuk karyawan
                    xdisc1 = 0
                    If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    If Len(VKary) = 9 Then
                        xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        
                        'simpan data kary ke promo_sales
                        ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                                Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                        ConnLocal.Execute StrSQL
                        If Linked Then
                            ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                            ConnServer.Execute StrSQL
                            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        End If
                    Else
                    xdisc1 = DiscNonKaryawan
                    End If
                Case 8 ' disc tambahan x% untuk karyawan
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                    If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    If Len(VKary) = 9 Then
                        .Recordset!ExtraDisc_pct = DiscNonKaryawan
                        .Recordset!ExtraDisc_amt = ((xqty * xharga) - (xqty * xharga * Pro_Disc / 100)) * (DiscNonKaryawan / 100)
                        
                        'simpan data kary ke promo_sales
                        ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                                Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                        ConnLocal.Execute StrSQL
                        If Linked Then
                            ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                            ConnServer.Execute StrSQL
                            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        End If
                        
                    End If
                Case 27 ' disc tambahan 17%
                    xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = TxtNom
                        .Recordset!ExtraDisc_amt = ((xqty * xharga) - Int((xqty * xharga * Pro_Disc / 100))) * (TxtNom / 100)
                'Case 9
                    'xdisc1 = Pro_Disc
                    'If xdisc1 = 100 Then GoTo loncat
                    'If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    'If Len(VKary) = 9 Then
                    '    xdisc1 = 10
                    '    .Recordset!ExtraDisc_pct = 0
                    '    .Recordset!ExtraDisc_amt = 0
                    '
                    '    'simpan data kary ke promo_sales
                    '    ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                    '    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                    '            Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                    '    ConnLocal.Execute StrSQL
                    '    If Linked Then
                    '        ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                    '        ConnServer.Execute StrSQL
                    '        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                    '    End If
                    
                    'End If

                Case 10, 24, 20
                    xdisc1 = Pro_Disc
                    If Pro_tipe = 20 Then
                    OnlyOncePromoTipe20 = True
                    End If
                    
                Case 12
                    xdisc1 = Pro_Disc
                    
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                    
                    If DiscMOP = 0 And xdisc1 > 0 Then
                        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                            DiscMOP = 1 'yes
                        Else
                            DiscMOP = 2 'no
                            VCekKartu = ""
                        End If
                    End If
                    
                    If DiscMOP = 1 Then
                        .Recordset!ExtraDisc_pct = 10
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * (xdisc1 / 100))) * 10 / 100
                    ElseIf DiscMOP = 2 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        VCekKartu = ""
                    End If
                Case 13
                    xdisc1 = .Recordset!Discount_Percentage
                    If DiscMOP = 0 Then
                        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                            DiscMOP = 1 'yes
                        Else
                            DiscMOP = 2 'no
                        End If
                    End If
                    
                    If DiscMOP = 1 Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - .Recordset!Discount_Amount) * Pro_Disc / 100
                    ElseIf DiscMOP = 2 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                Case Else
                    xdisc1 = 0
            End Select
loncat:
                
                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            If Pro_tipe = InfoPromox And xdisc1 <> 0 And CekPromoIDx = 1 Then
            HasilPromox = 1
            End If
            
            AdoLocal.Recordset.MoveNext
        Wend
        
        lblgrand_total = Format(vgtotal, "#,##0")
        'If Pro_Nm <> "" Then MsgBox Pro_Nm, vbInformation, "Oops.."
             
        End With
        If (Me.Caption = "SALES") Then
       'tipe 3,5,15
       
         With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void, Points_Received FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        
        If OnlyCheckPromo <> 2 Then
        OnlyCheckPromo = 0
        vgtotal = 0
        Tipe3Total = 0
        LimitTipe3 = 0
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe in ('3','5','15')  and Seqmentation = 0")
            If .Recordset!points_received = 0 Then
                     xdisc1 = .Recordset!Discount_Percentage
                    End If
            Select Case Pro_tipe
            
              Case 3 'promo karyawan
                    
                    If Mid(txtcard_no, 1, 5) = "CM999" And ScanApps = False Then
                    If Linked Then
                        RsAmbil.Open "Select b.ext1 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
                    Else
                        GoTo 1
                    End If
                    Bonus = RsAmbil!ext1
                    If Tipe3Total * (Pro_Disc / 100) >= RsAmbil!ext1 Then
                        VDiscBySTAR = RsAmbil!ext1
                    Else
                        VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
                    End If
                    
                    'If RsAmbil!ext1 - VDiscBySTAR > 0 Then
                    'VDiscBySTAR = RsAmbil!ext1
                    'End If
                    Nama_Promo = Pro_Nm
                    DiscStarProID = 3
                    RsAmbil.Close: Set RsAmbil = Nothing
                    End If
                 Case 5 'promo kartu
                 
                    Dim RsCekKartu As New ADODB.Recordset
                   
    
                    StrSQL = "Select * from promo_hdr where promo_id = '" & VCekKartu & "'"
                    RsCekKartu.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
                    VDiscBySTAR = (xqty * xharga - .Recordset!Discount_Amount) * Pro_Disc / 100
                    If RsCekKartu!islimit = 1 Then
                    isLimitCC = 1
                    If VDiscBySTAR > RsCekKartu!txt1 Then
                    VDiscBySTAR = RsCekKartu!txt1
                    If VDiscBySTAR / 1000 > (RsCekKartu!QtyLimit - RsCekKartu!QtyOut) Then
                    'MsgBox ("Promo Habis")
                    VDiscBySTAR = 0
                    isLimitCC = 0
                    DiscStarProID = 0
                    VCekKartu = ""
                    End If
                    End If
                    Else
                    isLimitCC = 2
                    End If
                    DiscStarProID = 5
                    RsCekKartu.Close: Set RsCekKartu = Nothing
                    
                'Case 9
                    
                '    VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
                '    Nama_Promo = Pro_Nm
                '    DiscStarProID = 9
                '    If VDiscBySTAR > LimitTipe9 Then
                '        VDiscBySTAR = LimitTipe9
                '    End If
                    
                Case 15
                    
                   
                    If Mid(txtcard_no, 1, 5) = "CM999" Then
                    If Linked Then
                        RsAmbil.Open "Select ISNULL(b.emp_whs,'') As emp_whs From Card a inner join List_Customer_Master_Member " & _
                        "b on a.Card_Nr = b.Card_Nr Where a.Card_Nr = '" & txtcard_no & "' and b.Emp_Whs = 'HO' or " & _
                        "(b.Emp_Whs = 'BO' And Emp_Title not in ('STA')) ", ConnServer, adOpenForwardOnly, adLockReadOnly
                    Else
                        GoTo 1
                    End If
                    
                    If Not RsAmbil.EOF Then
                    Else
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo 1
                    End If
                    If Trim(RsAmbil!emp_whs) = "" Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo 1
                    End If
                    
                    If .Recordset!points_received = 1 Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        GoTo 1
                    End If
                    
                    If xdisc_value > 0 Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * .Recordset!Discount_Percentage / 100)) * Pro_Disc / 100
                        .Recordset!Net_Price = xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    Else
                        xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                        .Recordset!Discount_Percentage = xdisc1
                        .Recordset!Discount_Amount = xdisc1_amt
                        .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    End If
                    Promo4Ever = True
                    RsAmbil.Close: Set RsAmbil = Nothing
                    Else
                    If .Recordset!points_received = 1 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                       .Recordset!points_received = 0
                    End If
                End Select
1:
If Pro_tipe = 15 And .Recordset!points_received = 1 Then
                Else
                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                End If
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            lblgrand_total = Format(vgtotal, "#,##0")
AdoLocal.Recordset.MoveNext
            Wend
                        
           


        Else
         OnlyCheckPromo = 0
         vgtotal = 0
         While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe in ('15')  and Seqmentation = 0")
            
            If .Recordset!points_received = 0 Then
                     xdisc1 = .Recordset!Discount_Percentage
                    End If
            Select Case Pro_tipe
                   
                Case 15
                    
                   
                    If Mid(txtcard_no, 1, 5) = "CM999" Then
                    If Linked Then
                        RsAmbil.Open "Select ISNULL(b.emp_whs,'') As emp_whs From Card a inner join List_Customer_Master_Member " & _
                        "b on a.Card_Nr = b.Card_Nr Where a.Card_Nr = '" & txtcard_no & "' and b.Emp_Whs = 'HO' or " & _
                        "(b.Emp_Whs = 'BO' And Emp_Title not in ('STA')) ", ConnServer, adOpenForwardOnly, adLockReadOnly
                    Else
                        GoTo loncat2
                    End If
                    
                    If Not RsAmbil.EOF Then
                    Else
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo loncat2
                    End If
                    If Trim(RsAmbil!emp_whs) = "" Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo loncat2
                    End If
                    
                    If .Recordset!points_received = 1 Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        GoTo loncat2
                    End If
                    
                    If xdisc_value > 0 Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * .Recordset!Discount_Percentage / 100)) * Pro_Disc / 100
                        .Recordset!Net_Price = xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    Else
                        xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                        .Recordset!Discount_Percentage = xdisc1
                        .Recordset!Discount_Amount = xdisc1_amt
                        .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    End If
                    Promo4Ever = True
                    RsAmbil.Close: Set RsAmbil = Nothing
                    Else
                    If .Recordset!points_received = 1 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                       .Recordset!points_received = 0
                    End If
                End Select
loncat2:
                If Pro_tipe = 15 And .Recordset!points_received = 1 Then
                Else
                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                End If
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            lblgrand_total = Format(vgtotal, "#,##0")
        AdoLocal.Recordset.MoveNext
            Wend
        End If
    
        End With
        
        
        'tambahan tipe 9 aja
        If (Me.Caption = "SALES") Then

       
         With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void, Points_Received FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        
       
        vgtotal = 0
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe in ('9')  and Seqmentation = 0")
            If .Recordset!points_received = 0 Then
                     xdisc1 = .Recordset!Discount_Percentage
                    End If
            Select Case Pro_tipe
            
              
                Case 9
                    
                    VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
                    Nama_Promo = Pro_Nm
                    DiscStarProID = 9
                    
                    
                
                End Select

                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            lblgrand_total = Format(vgtotal, "#,##0")
AdoLocal.Recordset.MoveNext
            Wend
                        
           

    
        End With
        End If
        
        If (VDiscBySTAR > 0) Then 'Or (Me.Caption = "REFUND" And VDiscBySTAR < 0) Then
        Select Case DiscStarProID
        
        Case 3
        If Linked Then
                        
                        Sisa_Bonus = Bonus - VDiscBySTAR
                        ConnLocal.Execute "Update b set b.ext1 = " & Sisa_Bonus & " From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                            ConnServer.Execute "Update b set b.ext1 = " & Sisa_Bonus & " From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                             ConnLocal.Execute "Delete from paid where transaction_number='" & vno_trans & "'"
                    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
                    "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
                    vno_trans & "','31','1','IDR','1','','" & Nama_Promo & "'," & VDiscBySTAR & ",'" & VShift & "')"
        End If
        Case 9
        If Linked Then
                        If VDiscBySTAR > LimitTipe9 Then
                            VDiscBySTAR = LimitTipe9
                        End If
                       
                        ConnLocal.Execute "Update b set b.ext1 = 0 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                            ConnServer.Execute "Update b set b.ext1 = 0 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                             ConnLocal.Execute "Delete from paid where transaction_number='" & vno_trans & "'"
                    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
                    "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
                    vno_trans & "','31','1','IDR','1','','" & Nama_Promo & "'," & VDiscBySTAR & ",'" & VShift & "')"
        End If
        Case 5
                    If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                    ConnLocal.Execute "Delete from paid where transaction_number='" & vno_trans & "'"
                    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
                    "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
                    vno_trans & "','31','1','IDR','1','','" & Nama_Promo & "'," & VDiscBySTAR & ",'" & VShift & "')"
                    Else
                            VDiscBySTAR = 0
                            VCekKartu = ""
                    End If
                    
        End Select
        
        End If
        Else 'refund
         With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void, Points_Received FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        vgtotal = 0
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe in ('15')  and Seqmentation = 0")
            If .Recordset!points_received = 0 Then
                xdisc1 = .Recordset!Discount_Percentage
            End If
         Select Case Pro_tipe
         Case 15
                    
                   
                    If Mid(txtcard_no, 1, 5) = "CM999" Then
                    If Linked Then
                        RsAmbil.Open "Select ISNULL(b.emp_whs,'') As emp_whs From Card a inner join List_Customer_Master_Member " & _
                        "b on a.Card_Nr = b.Card_Nr Where a.Card_Nr = '" & txtcard_no & "' and b.Emp_Whs = 'HO' or " & _
                        "(b.Emp_Whs = 'BO' And Emp_Title not in ('STA')) ", ConnServer, adOpenForwardOnly, adLockReadOnly
                    Else
                        GoTo 3
                    End If
                    
                    If Not RsAmbil.EOF Then
                    Else
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo 3
                    End If
                    If Trim(RsAmbil!emp_whs) = "" Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        .Recordset!points_received = 0
                        GoTo 3
                    End If
                    
                    If .Recordset!points_received = 1 Then
                        RsAmbil.Close: Set RsAmbil = Nothing
                        GoTo 3
                    End If
                    
                    If xdisc_value < 0 Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * .Recordset!Discount_Percentage / 100)) * Pro_Disc / 100
                        .Recordset!Net_Price = xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    Else
                        xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                        .Recordset!Discount_Percentage = xdisc1
                        .Recordset!Discount_Amount = xdisc1_amt
                        .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                        .Recordset!points_received = 1
                    End If
                    Promo4Ever = True
                    RsAmbil.Close: Set RsAmbil = Nothing
                    Else
                    If .Recordset!points_received = 1 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                       .Recordset!points_received = 0
                    End If
         End Select
3:
If Pro_tipe = 15 And .Recordset!points_received = 1 Then
                Else
                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                End If
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            lblgrand_total = Format(vgtotal, "#,##0")
AdoLocal.Recordset.MoveNext
        
        Wend
        End With
        End If
        
End Sub

Private Sub Cek_Promo_Seq(ByVal Promoid As String)
Dim RsPromo As New ADODB.Recordset, RsAmbil As New ADODB.Recordset, RsKaryawan As New ADODB.Recordset
Dim xqty As Integer, xharga As Long, xdisc1 As Byte, xdisc1_amt As Long
Dim Pro_tipe As Byte, Pro_Nm As String, Pro_Disc As Byte, Nama_Promo As String
Dim DiscMOP As Byte, Sisa_Bonus As Long, Bonus As Long
        ConnLocal.Execute "update Sales_Transaction_Details set Discount_Amount = 0,Discount_Percentage = 0,ExtraDisc_Amt = 0,ExtraDisc_Pct = 0 where transaction_number='" & vno_trans & "' and Points_Received in (1)"
        ConnLocal.Execute "Update Sales_Transaction_Details set Points_Received = 0 where transaction_number='" & vno_trans & "' and Points_Received in (1)"
        With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void, Points_Received  FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        
        vgtotal = 0
        VDiscBySTAR = 0
        DiscMOP = 0
        VCekKartu = ""
        Bonus = 0
        
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe not in (3,5,15)  and Seqmentation <> 0 and ph.promo_id= '" & Promoid & "'")
            
            Select Case Pro_tipe
                Case 0 'tidak ada promo
                    UpdateStatusSeqDetail = False
                    xdisc1 = .Recordset!Discount_Percentage
                Case 2, 11, 14, 17, 18, 19, 26, 28
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                Case 22, 23, 25
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                Case 21
                If txtcard_no <> "CM000-00000" Then
                    .Recordset!ExtraDisc_pct = Pro_Disc
                    .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (Pro_Disc / 100)
                Else
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                End If
'                Case 3 Sebelum 6/6/2016 utuk disc karyawan
'                    xdisc1 = .Recordset!Discount_Percentage
'                    VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
'                    Nama_Promo = Pro_Nm
                Case 4
                    xdisc1 = Pro_Disc
                    If txtcard_no <> "CM000-00000" Then
                        .Recordset!ExtraDisc_pct = IIf(Not IsNumeric(TxtNom), 0, TxtNom)
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * Pro_Disc / 100)) * IIf(Not IsNumeric(TxtNom), 0, TxtNom) / 100
                    Else
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                
                    
                 Case 29
                    xdisc1 = .Recordset!Discount_Percentage
                    If txtcard_no <> "CM000-00000" Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (Pro_Disc / 100)
                    Else
                        .Recordset!ExtraDisc_pct = IIf(Not IsNumeric(TxtNom), 0, TxtNom)
                        .Recordset!ExtraDisc_amt = (xqty * xharga - Int((xqty * xharga) * (xdisc1 / 100))) * (IIf(Not IsNumeric(TxtNom), 0, TxtNom) / 100)
                    End If
                'Case 5
                '    xdisc1 = .Recordset!Discount_Percentage
                '    If DiscMOP = 0 Then
                '        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                '            DiscMOP = 1 'yes
                '        Else
                '            DiscMOP = 2 'no
                '            VCekKartu = ""
                '        End If
                '    End If
                    
                '    If DiscMOP = 1 Then
                '        .Recordset!ExtraDisc_pct = Pro_Disc
                '        .Recordset!ExtraDisc_amt = (xqty * xharga - .Recordset!Discount_Amount) * Pro_Disc / 100
                '    ElseIf DiscMOP = 2 Then
                '        .Recordset!ExtraDisc_pct = 0
                '        .Recordset!ExtraDisc_amt = 0
                '        VCekKartu = ""
                '        VDiscBySTAR = 0
                '    End If
                Case 6
                    xdisc1 = Pro_Disc
                Case 7 ' disc 10% untuk karyawan
                    xdisc1 = 0
                    If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    If Len(VKary) = 9 Then
                        xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        
                        'simpan data kary ke promo_sales
                        ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                                Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                        ConnLocal.Execute StrSQL
                        If Linked Then
                            ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                            ConnServer.Execute StrSQL
                            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        End If
                    Else
                    xdisc1 = DiscNonKaryawan
                    End If
                Case 8 ' disc tambahan x% untuk karyawan
                    xdisc1 = Pro_Disc
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                    If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    If Len(VKary) = 9 Then
                        .Recordset!ExtraDisc_pct = DiscNonKaryawan
                        .Recordset!ExtraDisc_amt = ((xqty * xharga) - (xqty * xharga * Pro_Disc / 100)) * (DiscNonKaryawan / 100)
                        
                        'simpan data kary ke promo_sales
                        ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                                Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                        ConnLocal.Execute StrSQL
                        If Linked Then
                            ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                            ConnServer.Execute StrSQL
                            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        End If
                        
                    End If
                Case 27 ' disc tambahan 17%
                    xdisc1 = Pro_Disc
                        .Recordset!ExtraDisc_pct = TxtNom
                        .Recordset!ExtraDisc_amt = ((xqty * xharga) - Int((xqty * xharga * Pro_Disc / 100))) * (TxtNom / 100)
                Case 9
                    xdisc1 = Pro_Disc
                    If xdisc1 = 100 Then GoTo loncat
                    If VKary = "" And VTanya = False Then frmkaryawan.Show 1
                    If Len(VKary) = 9 Then
                        xdisc1 = 10
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        
                        'simpan data kary ke promo_sales
                        ConnLocal.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                                Pro_tipe & "', '" & vno_trans & "', " & VKary & ", " & .Recordset!Seq & ", '00')"
    
                        ConnLocal.Execute StrSQL
                        If Linked Then
                            ConnServer.Execute "delete promo_sales where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                            ConnServer.Execute StrSQL
                            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & Pro_tipe & "' and transaction_number='" & vno_trans & "'"
                        End If
                    
                    End If

                Case 10, 24, 20
                    xdisc1 = Pro_Disc
                    If Pro_tipe = 20 Then
                    OnlyOncePromoTipe20 = True
                    End If
                    
                Case 12
                    xdisc1 = Pro_Disc
                    
                    .Recordset!ExtraDisc_pct = 0
                    .Recordset!ExtraDisc_amt = 0
                    
                    If DiscMOP = 0 And xdisc1 > 0 Then
                        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                            DiscMOP = 1 'yes
                        Else
                            DiscMOP = 2 'no
                            VCekKartu = ""
                        End If
                    End If
                    
                    If DiscMOP = 1 Then
                        .Recordset!ExtraDisc_pct = 10
                        .Recordset!ExtraDisc_amt = (xqty * xharga - (xqty * xharga * (xdisc1 / 100))) * 10 / 100
                    ElseIf DiscMOP = 2 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                        VCekKartu = ""
                    End If
                Case 13
                    xdisc1 = .Recordset!Discount_Percentage
                    If DiscMOP = 0 Then
                        If MsgBox("Apakah customer menggunakan " & Pro_Nm & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                            DiscMOP = 1 'yes
                        Else
                            DiscMOP = 2 'no
                        End If
                    End If
                    
                    If DiscMOP = 1 Then
                        .Recordset!ExtraDisc_pct = Pro_Disc
                        .Recordset!ExtraDisc_amt = (xqty * xharga - .Recordset!Discount_Amount) * Pro_Disc / 100
                    ElseIf DiscMOP = 2 Then
                        .Recordset!ExtraDisc_pct = 0
                        .Recordset!ExtraDisc_amt = 0
                    End If
                Case Else
                    xdisc1 = 0
            End Select
loncat:
                
                xdisc1_amt = xqty * xharga * (xdisc1 / 100)
                    .Recordset!Discount_Percentage = xdisc1
                    .Recordset!Discount_Amount = xdisc1_amt
                    .Recordset!Net_Price = xqty * xharga - xdisc1_amt - .Recordset!ExtraDisc_amt
                    .Recordset!points_received = 2
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            If Pro_tipe = InfoPromox And xdisc1 <> 0 And CekPromoIDx = 1 Then
            HasilPromox = 1
            End If
            If UpdateStatusSeqDetail = True Then
                    SeqCountInt = SeqCountInt + 1
                    SeqCount(SeqCountInt) = Promoid
            End If
            AdoLocal.Recordset.MoveNext
        Wend
        
        lblgrand_total = Format(vgtotal, "#,##0")
        'If Pro_Nm <> "" Then MsgBox Pro_Nm, vbInformation, "Oops.."
             
        End With
        
        'promo seq tipe 3
        With AdoLocal
        .RecordSource = "SELECT Seq, Flag_Paket_Discount, PLU, Item_Description, Price, Qty, Discount_Percentage, " & _
                        "Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, flag_void, Points_Received FROM Sales_Transaction_Details " & _
                        "where transaction_number='" & vno_trans & "'"
        .Refresh
        vgtotal = 0
        UpdateStatusSeqDetail = True
        While Not .Recordset.EOF
            xqty = .Recordset!Qty
            xharga = .Recordset!price
            Call Cari_Promo(.Recordset!plu, vno_trans, Pro_tipe, Pro_Nm, Pro_Disc, .Recordset!Seq, " and tipe in ('3')  and Seqmentation <> 0")
            Select Case Pro_tipe
                Case 0 'tidak ada promo
                    UpdateStatusSeqDetail = False
                Case 3 'promo disc STAR
                If LimitTipe3 = 0 Then
                    VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
                    DiscStarProID = 3
                Else
                    If Tipe3Total * (Pro_Disc / 100) >= LimitTipe3 Then
                        VDiscBySTAR = LimitTipe3
                    Else
                        VDiscBySTAR = VDiscBySTAR + ((xqty * xharga - .Recordset!Discount_Amount - .Recordset!ExtraDisc_amt) * Pro_Disc / 100)
                    End If
                    Nama_Promo = Pro_Nm
                    DiscStarProID = 3
                End If
                 
                End Select
                If UpdateStatusSeqDetail = True Then
                    SeqCountInt = SeqCountInt + 1
                    SeqCount(SeqCountInt) = Promoid
                End If
            .Recordset.Update
            
            vgtotal = vgtotal + .Recordset!Net_Price
            lblgrand_total = Format(vgtotal, "#,##0")
            AdoLocal.Recordset.MoveNext
            Wend
        End With
        
        If (VDiscBySTAR > 0) Then
        Select Case DiscStarProID
        Case 3
        If Linked Then
                    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
                    "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
                    vno_trans & "','31','1','IDR','1','','" & Nama_Promo & "'," & VDiscBySTAR & ",'" & VShift & "')"
        End If
        End Select
        
        End If
End Sub

Private Sub Cari_Promo(codebar As String, No_trans As String, ByRef Promo_Tipe As Byte, ByRef Promo_Nm As String, ByRef Promo_Disc As Byte, Seq As Byte, ByRef ExPromo As String)
Dim RsPromo As New ADODB.Recordset, RsHitung As New ADODB.Recordset, RsAmbil As New ADODB.Recordset
Dim diskon As Byte, jml As Byte, min_belanja As Long
    
    diskon = 0
    Promo_Tipe = 0
    Promo_Nm = ""
    Promo_Disc = 0
    
    RsPromo.Open "select ph.promo_id, promo_name, min_purchase, min_member, disc, tipe, txt1, txt2 from promo_hdr ph inner join " & _
                 "promo_dtl pd on ph.promo_id=pd.promo_id where getdate() Between Start_Date And End_Date and aktif=1 " & _
                 "and tipe <30 " & ExPromo & " and  PLU ='" & codebar & "' order by tipe desc", ConnLocal, adOpenForwardOnly, adLockReadOnly
    
    If RsPromo.EOF Then
        RsPromo.Close: Set RsPromo = Nothing
        Exit Sub
    End If
            
    While Not RsPromo.EOF
    
        If Left(Star_Id, 6) = "100000" Or Star_Id = "" Then
            min_belanja = RsPromo!min_purchase
        Else
            min_belanja = RsPromo!min_member
        End If
        If PromoIDx = RsPromo!promo_id Then
        CekPromoIDx = 1
        End If
        Select Case RsPromo!Tipe
        Case 2, 7, 8, 19, 21, 26 'disc progressive mis : min 100 disc 10, min 200 disc 20
            If RsPromo!disc > diskon Then
                If min_belanja > 0 Then
                    RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                                  "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", _
                                  ConnLocal, adOpenForwardOnly, adLockReadOnly
                    If RsHitung!Berapa > 0 Then 'sales
                        If RsHitung!Berapa >= min_belanja Then
                            diskon = RsPromo!disc
                        End If
                    Else 'refund
                        If RsHitung!Berapa <= (-1 * min_belanja) Then
                            diskon = RsPromo!disc
                        End If
                    End If
                    Promo_Nm = RsPromo!promo_name
                    Promo_Tipe = RsPromo!Tipe
                    RsHitung.Close: Set RsHitung = Nothing
                Else
                    diskon = RsPromo!disc
                    Promo_Tipe = RsPromo!Tipe
                    Promo_Nm = RsPromo!promo_name
                End If
            End If
            Promo_Disc = diskon
            If RsPromo!Tipe = 26 Then
                If VIsKKG = False Then Promo_Disc = 0
            End If
            DiscNonKaryawan = 0
            If RsPromo!Tipe = 7 Or RsPromo!Tipe = 8 Then
                 DiscNonKaryawan = IIf(Not IsNumeric(RsPromo!txt1), 0, RsPromo!txt1)
            End If
            If RsPromo!Tipe = 19 Then
                 RsHitung.Open "select sum(qty) as Berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                                  "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!txt1 & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            If Not RsHitung.EOF Then
                Promo_Disc = 0
                If RsHitung!Berapa > 0 Then Promo_Disc = diskon
            Else
                Promo_Disc = 0
            End If
            RsHitung.Close: Set RsHitung = Nothing
            End If
            If RsPromo!Tipe = 21 Then
                 RsHitung.Open "Select b.Cust_Gender From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "' And b.Cust_Gender = '" & RsPromo!txt1 & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            If Not RsHitung.EOF Then
                Promo_Disc = diskon
            Else
                Promo_Disc = 0
            End If
            RsHitung.Close: Set RsHitung = Nothing
            End If
        Case 3, 9, 4, 5, 13, 29, 27
        '3 = disc by STAR
        '4 = disc for mystar card holder
        '5 = disc by credit card/payment
        If RsPromo!Tipe = 3 Or RsPromo!Tipe = 9 Then
            RsHitung.Open "select SUM(qty*price-Discount_Amount-ExtraDisc_Amt) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            Tipe3Total = RsHitung!Berapa
            LimitTipe3 = 0
            LimitTipe9 = 0
            If RsPromo!txt1 <> "" Then
                LimitTipe3 = RsPromo!txt1
                LimitTipe9 = RsPromo!txt1
            End If
        Else
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
        End If
        
        If RsPromo!Tipe = 9 Then
        If Mid(VReg_ID, 1, 1) = Right(RsPromo!txt2, 1) Then
            Exit Sub
        End If
            If Left(RsPromo!txt2, 1) = 1 Then
              If Mid(txtcard_no, 1, 5) = "CM000" Or Mid(txtcard_no, 1, 5) = "CM999" Then
                  Exit Sub
              End If
              If Linked Then
                  RsAmbil.Open "Select b.ext1 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
                  If RsAmbil!ext1 = 0 Then
                        Exit Sub
                  End If
              Else
                  Exit Sub
              End If
            End If
        End If
        
        
            If RsHitung!Berapa > 0 Then 'sales
                If RsHitung!Berapa >= min_belanja Then
                    Promo_Tipe = RsPromo!Tipe
                    Promo_Nm = RsPromo!promo_name
                    Promo_Disc = RsPromo!disc
                    If RsPromo!Tipe = 5 Then VCekKartu = RsPromo!promo_id
                End If
            Else 'refund
                If RsHitung!Berapa <= (-1 * min_belanja) Then
                    Promo_Tipe = RsPromo!Tipe
                    Promo_Nm = RsPromo!promo_name
                    Promo_Disc = RsPromo!disc
                End If
            End If
           
            If RsPromo!Tipe = 29 Or RsPromo!Tipe = 4 Or RsPromo!Tipe = 27 Then
            TxtNom = IIf(Not IsNumeric(RsPromo!txt1), 0, RsPromo!txt1)
            End If
            
            RsHitung.Close: Set RsHitung = Nothing
                        Exit Sub
            
        Case 15
        xdisc_value = 0
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "' ", ConnLocal, adOpenForwardOnly, adLockReadOnly

            If RsHitung!Berapa > 0 Then 'sales
                If RsHitung!Berapa >= min_belanja Then
                    Promo_Tipe = RsPromo!Tipe
                    Promo_Nm = RsPromo!promo_name
                    RsHitung.Close
                    RsHitung.Open "select SUM(qty*price) as berapa,SUM(discount_amount) as disc_value from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'  and  sd.PLU ='" & codebar & "' And Seq = '" & Seq & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                    xdisc_value = RsHitung!disc_value
                    If xdisc_value > 0 Then
                    Promo_Disc = RsPromo!txt1
                    Else
                    Promo_Disc = RsPromo!disc
                    End If
                End If
            Else 'refund
                If RsHitung!Berapa <= (-1 * min_belanja) Then
                    Promo_Tipe = RsPromo!Tipe
                    Promo_Nm = RsPromo!promo_name
                    RsHitung.Close
                    RsHitung.Open "select SUM(qty*price) as berapa,SUM(discount_amount) as disc_value from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'  and  sd.PLU ='" & codebar & "' And Seq = '" & Seq & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                    xdisc_value = RsHitung!disc_value
                    If xdisc_value < 0 Then
                    Promo_Disc = RsPromo!txt1
                    Else
                    Promo_Disc = RsPromo!disc
                    End If
                End If
            End If
            
            
            RsHitung.Close: Set RsHitung = Nothing
            Exit Sub

        Case 6 'disc #% untuk item kedua yang termurah
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            If RsHitung!Berapa >= 0 Then  'sales
            
            RsHitung.Close
            RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                          
            jml = roundDown(RsHitung.RecordCount / 2)
            
            StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                     "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "' order by amount, net_price ) aa " & _
                     "where PLU = '" & codebar & "' and seq='" & Seq & "'"
            Else 'refund
                    RsHitung.Close
                    RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                                  "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                                  
                    jml = roundDown(RsHitung.RecordCount / 2)
                    
                    StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                             "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "' order by amount desc, net_price desc) aa " & _
                             "where PLU = '" & codebar & "' and seq='" & Seq & "'"
                    End If
            RsHitung.Close
            RsHitung.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name

            If Not RsHitung.EOF Then
                Promo_Disc = RsPromo!disc
            Else
                Promo_Disc = 0
            End If

            RsHitung.Close: Set RsHitung = Nothing
        Case 10  'disc #% untuk item ketiga yang termurah
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            If RsHitung!Berapa > 0 Then 'sales
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = roundDown(RsHitung.RecordCount / 3)
                
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "' order by amount, net_price ) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
                Else 'refund
                
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = roundDown(RsHitung.RecordCount / 3)
                
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "' order by amount desc, net_price desc) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
            End If
            
            RsHitung.Close
            RsHitung.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name

            If Not RsHitung.EOF Then
                Promo_Disc = RsPromo!disc
            Else
                Promo_Disc = 0
            End If

            RsHitung.Close: Set RsHitung = Nothing
        Case 11, 12, 14 '11 = disc #% untuk pembelian 2 pcs
            '12 = disc #% untuk pembelian 2 pcs + kartu kredit addtional #%
            '14 = disc #% pakai kupon disc jadi disc 30% (Loreal)
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

            If RsHitung!Berapa >= 2 Or RsHitung!Berapa <= -2 Then
                Promo_Tipe = RsPromo!Tipe
                Promo_Nm = RsPromo!promo_name
                If RsPromo!Tipe = 14 Then
                    If MsgBox("Apakah customer menggunakan kupon disc?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        Promo_Disc = 30 'yes
                    Else
                        Promo_Disc = 0 'no
                        If min_belanja > 0 Then
                            Dim rshitmin As New ADODB.Recordset
                            rshitmin.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", _
                                          ConnLocal, adOpenForwardOnly, adLockReadOnly
                    
                                If rshitmin!Berapa > 0 Then 'sales
                                    If rshitmin!Berapa >= min_belanja Then
                                        Promo_Disc = RsPromo!disc
                                    End If
                                Else 'refund
                                    If rshitmin!Berapa <= (-1 * min_belanja) Then
                                        Promo_Disc = RsPromo!disc
                                    End If
                                End If
                            rshitmin.Close: Set rshitmin = Nothing
                        Else
                            Promo_Disc = RsPromo!disc 'no
                        End If
                    End If
                Else
                    Promo_Disc = RsPromo!disc
                End If
                If RsPromo!Tipe = 12 Then VCekKartu = RsPromo!promo_id
            Else
                Promo_Tipe = RsPromo!Tipe
                Promo_Nm = RsPromo!promo_name
                Promo_Disc = 0
            End If
            RsHitung.Close: Set RsHitung = Nothing
            Exit Sub
        Case 17  '= disc progressif untuk pembelian 2 pcs dan 3 pcs
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

            'If RsPromo!tipe = 15 Then
            '    If RsHitung!Berapa >= 2 Then
            '       Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa >= 3 Then Promo_Disc = RsPromo!disc + 5
            '    ElseIf RsHitung!Berapa <= -2 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa <= -3 Then Promo_Disc = RsPromo!disc + 5
            '    Else
            '        Promo_Disc = 0
            '    End If
            'ElseIf RsPromo!tipe = 16 Then
            '    If RsHitung!Berapa >= 2 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa >= 3 Then Promo_Disc = RsPromo!disc + 10
            '    ElseIf RsHitung!Berapa <= -2 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa <= -3 Then Promo_Disc = RsPromo!disc + 10
            '    Else
            '        Promo_Disc = 0
            '    End If
            ' Else
            If RsHitung!Berapa >= 2 Then
                Promo_Disc = RsPromo!disc
                If RsHitung!Berapa >= 3 Then Promo_Disc = RsPromo!txt1
            ElseIf RsHitung!Berapa <= -2 Then
                Promo_Disc = RsPromo!disc
                If RsHitung!Berapa <= -3 Then Promo_Disc = RsPromo!txt1
            Else
                Promo_Disc = 0
            End If
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name
            RsHitung.Close: Set RsHitung = Nothing
        Case 18, 28  '= disc progressif untuk pembelian 1 pcs dan 2 pcs, (note tipe 19 dan 20 dirubah)
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

            If RsPromo!Tipe = 18 Then
                If RsHitung!Berapa >= 1 Then
                    Promo_Disc = RsPromo!disc
                    If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!txt1
                ElseIf RsHitung!Berapa <= -1 Then
                    Promo_Disc = RsPromo!disc
                    If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!txt1
                Else
                    Promo_Disc = 0
                End If
            'ElseIf RsPromo!tipe = 19 Then
            '    If RsHitung!Berapa >= 1 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!disc + 10
            '    ElseIf RsHitung!Berapa <= -1 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!disc + 10
            '    Else
            '        Promo_Disc = 0
            '    End If
            ' ElseIf RsPromo!tipe = 20 Then
            '    If RsHitung!Berapa >= 1 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!disc + 20
            '   ElseIf RsHitung!Berapa <= -1 Then
            '        Promo_Disc = RsPromo!disc
            '       If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!disc + 20
            '    Else
            '        Promo_Disc = 0
            '   End If
            'ElseIf RsPromo!tipe = 21 Then (Dirubah tipe promo baru 16/11/2016 Disc MSC Female)
            '    If RsHitung!Berapa >= 1 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!disc + 28
            '   ElseIf RsHitung!Berapa <= -1 Then
            '        Promo_Disc = RsPromo!disc
            '        If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!disc + 28
            '    Else
            '       Promo_Disc = 0
            '    End If
            ElseIf RsPromo!Tipe = 28 Then
                If RsHitung!Berapa >= 1 Then
                    Promo_Disc = RsPromo!disc
                    If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!disc + 20
                ElseIf RsHitung!Berapa <= -1 Then
                    Promo_Disc = RsPromo!disc
                    If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!disc + 20
                Else
                    Promo_Disc = 0
                End If
            End If
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name
            RsHitung.Close: Set RsHitung = Nothing
        Case 22 '= disc progressif untuk pembelian 1, 2, 3 pcs
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            If RsHitung!Berapa >= 1 Then
                Promo_Disc = RsPromo!disc
                If RsHitung!Berapa >= 2 Then Promo_Disc = RsPromo!txt1
                If RsHitung!Berapa >= 3 Then Promo_Disc = RsPromo!txt2
            ElseIf RsHitung!Berapa <= -1 Then
                Promo_Disc = RsPromo!disc
                If RsHitung!Berapa <= -2 Then Promo_Disc = RsPromo!txt1
                If RsHitung!Berapa <= -3 Then Promo_Disc = RsPromo!txt2
            Else
                Promo_Disc = 0
            End If
                       
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name
            RsHitung.Close: Set RsHitung = Nothing
       Case 23 '23 = disc #% untuk min pembelian 3 pcs
            Dim diskon23 As Byte
            diskon23 = 0
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly

            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name
            'If RsHitung!Berapa >= 3 Or RsHitung!Berapa <= -3 Then
            '    Promo_Disc = RsPromo!disc
            'Else
            '    Promo_Disc = 0
            'End If
            If RsHitung!Berapa >= 2 Then
                diskon23 = RsPromo!disc
                If RsHitung!Berapa >= 3 Then diskon23 = RsPromo!txt1
            ElseIf RsHitung!Berapa <= -2 Then
                diskon23 = RsPromo!disc
                If RsHitung!Berapa <= -3 Then diskon23 = RsPromo!txt1
            Else
                diskon23 = 0
            End If
            RsHitung.Close: Set RsHitung = Nothing
            RsHitung.Open "select SUM(qty) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                                  "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!txt2 & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            If Not RsHitung.EOF Then
                Promo_Disc = 0
                If RsHitung!Berapa > 0 Then Promo_Disc = diskon23
            Else
                Promo_Disc = 0
            End If
            RsHitung.Close: Set RsHitung = Nothing
        Case 24 'beli 3 gratis 2
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            If RsHitung!Berapa > 0 Then 'sales
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = 2 * roundDown(RsHitung.RecordCount / 5)
                
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "' order by amount, net_price, seq ) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
                Else 'refund
                
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = 2 * roundDown(RsHitung.RecordCount / 5)
                
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "' order by amount desc, net_price, seq desc) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
            End If
            
            RsHitung.Close
            RsHitung.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name

            If Not RsHitung.EOF Then
                Promo_Disc = RsPromo!disc
            Else
                Promo_Disc = 0
            End If

            RsHitung.Close: Set RsHitung = Nothing
        Case 20 'beli n gratis n kondisi txt1
            RsHitung.Open "select SUM(qty*price) as berapa from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                          "where Transaction_Number='" & No_trans & "' and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            If RsHitung!Berapa > 0 Then 'sales
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = Right(RsPromo!txt1, 1) * roundDown(RsHitung.RecordCount / (Int(Right(RsPromo!txt1, 1)) + Int(Left(RsPromo!txt1, 1))))

                
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=1 and promo_id='" & RsPromo!promo_id & "' order by amount, net_price ) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
            Else 'refund
                
                RsHitung.Close
                RsHitung.Open "select * from Sales_Transaction_Details sd inner join PROMO_DTL pd on sd.PLU = pd.PLU " & _
                              "where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                              
                jml = Right(RsPromo!txt1, 1) * roundDown(RsHitung.RecordCount / (Int(Right(RsPromo!txt1, 1)) + Int(Left(RsPromo!txt1, 1))))
                StrSQL = "select * from(select top " & jml & " sd.seq, sd.plu from  Sales_Transaction_Details sd inner join PROMO_DTL pd " & _
                         "on sd.PLU = pd.PLU where Transaction_Number='" & No_trans & "' and flag_void=0 and qty=-1 and promo_id='" & RsPromo!promo_id & "' order by amount desc, net_price desc) aa " & _
                         "where PLU = '" & codebar & "' and seq='" & Seq & "'"
            End If
            
            RsHitung.Close
            RsHitung.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
            
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name

            If Not RsHitung.EOF Then
                Promo_Disc = RsPromo!disc
            Else
                Promo_Disc = 0
            End If

            RsHitung.Close: Set RsHitung = Nothing
       Case 25 'disc 10% untuk member membeli voucher
            StrSQL = "select * from promo_sales where " & _
                     "Transaction_Number='" & Star_No & "' and promo_id='" & RsPromo!promo_id & "'"
                    
            If Linked Then
                RsHitung.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
            Else
                RsHitung.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
            End If
                    
            Promo_Tipe = RsPromo!Tipe
            Promo_Nm = RsPromo!promo_name
            
            If Not RsHitung.EOF Then
                Promo_Disc = 0
                MsgBox "Customer  " & Left(Star_No, 11) & " sudah pernah mendapatkan discount 10%  ", vbOKOnly + vbInformation, "Oops.."
            Else
                If lblgrand_total <= 2000000 Then
                    Promo_Disc = RsPromo!disc
                Else
                    Promo_Disc = 0
                    MsgBox "Maksimal 2 juta rupiah", vbOKOnly + vbInformation, "Oops.."
                End If
            End If
            
            RsHitung.Close: Set RsHitung = Nothing
        End Select
        RsPromo.MoveNext
    Wend
    RsPromo.Close: Set RsPromo = Nothing
End Sub
