VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{CCB90040-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSCashDrawer.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POINT OF SALES"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin TDBText6Ctl.TDBText txtinfo 
      Height          =   1650
      Left            =   150
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   300
      Width           =   6615
      _Version        =   65536
      _ExtentX        =   11668
      _ExtentY        =   2910
      Caption         =   "frmMenu.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMenu.frx":0936
      Key             =   "frmMenu.frx":0954
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
      MultiLine       =   -1
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TrueOleDBGrid80.TDBGrid Grid_Hold 
      Bindings        =   "frmMenu.frx":0998
      Height          =   5610
      Left            =   6900
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   300
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9895
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO"
      Columns(0).DataField=   "nomor"
      Columns(0).DataWidth=   11
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NILAI"
      Columns(1).DataField=   "Net_amount"
      Columns(1).DataWidth=   22
      Columns(1).NumberFormat=   "#,##0"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Transaction_Number"
      Columns(2).DataField=   "Transaction_Number"
      Columns(2).DataWidth=   21
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2805"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2672"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._AlignLeft=0"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=238"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=106"
      Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      Enabled         =   0   'False
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
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
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   150
      TabIndex        =   11
      Top             =   2280
      Width           =   6615
      Begin VB.CommandButton cmdMenu 
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
         Height          =   1000
         Index           =   8
         Left            =   5250
         Picture         =   "frmMenu.frx":09AF
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   9
         Left            =   5250
         Picture         =   "frmMenu.frx":13B1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "OPEN DRAWER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   3
         Left            =   3975
         Picture         =   "frmMenu.frx":1DB3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "RELEASE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   7
         Left            =   3975
         Picture         =   "frmMenu.frx":27B5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "REPRINT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   6
         Left            =   2700
         Picture         =   "frmMenu.frx":31B7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "REFUND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   5
         Left            =   1425
         Picture         =   "frmMenu.frx":3BB9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   4
         Left            =   150
         Picture         =   "frmMenu.frx":45BB
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "CLOSE REGISTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   2
         Left            =   2700
         Picture         =   "frmMenu.frx":4FBD
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "CLOSE SHIFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   1
         Left            =   1425
         Picture         =   "frmMenu.frx":59BF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1250
      End
      Begin VB.CommandButton cmdMenu 
         Caption         =   "CASH OPEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   0
         Left            =   150
         Picture         =   "frmMenu.frx":63C1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1250
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   180
      TabIndex        =   10
      Top             =   4920
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3300
         Top             =   600
      End
      Begin VB.Label lblline 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3450
         TabIndex        =   22
         Top             =   675
         Width           =   3090
      End
      Begin VB.Label lbljam 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3450
         TabIndex        =   17
         Top             =   375
         Width           =   3090
      End
      Begin VB.Label lbltgl 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3450
         TabIndex        =   16
         Top             =   75
         Width           =   3090
      End
      Begin VB.Label lblkasir 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   675
         Width           =   3090
      End
      Begin VB.Label lblreg 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   375
         Width           =   3090
      End
      Begin VB.Label lblbranch 
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   75
         Width           =   3090
      End
   End
   Begin MSAdodcLib.Adodc AdoLocal 
      Height          =   330
      Left            =   6975
      Top             =   5400
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CEK BIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   200
      TabIndex        =   24
      Top             =   75
      Width           =   900
   End
   Begin OposCashDrawer_CCOCtl.OPOSCashDrawer OPOSCashDrawer1 
      Left            =   75
      OleObjectBlob   =   "frmMenu.frx":6DC3
      Top             =   4425
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INFO REGISTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4575
      TabIndex        =   21
      Top             =   4680
      Width           =   2190
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4575
      TabIndex        =   20
      Top             =   2040
      Width           =   2190
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI PENDING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8325
      TabIndex        =   19
      Top             =   75
      Width           =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INFO PROMO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      TabIndex        =   18
      Top             =   75
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMenu_Click(Index As Integer)
Dim RsCek As New ADODB.Recordset
    Select Case Index
    Case 0 'CASH OPEN
        If cmdMenu(0).Caption = "CASH OPEN" Then
            frmCashOpen.Show 1
        Else
            frmPassword.Show 1
        End If
    Case 1 'CLOSE SHIFT
        If Cfg_Get("Device", "X_ReadPrint", App.Path & "\config.ini") = 1 Then
            If Not Super(2) Then Exit Sub
        Else
            If Linked Then
                RsCek.Open "Select Spending_Program_ID As Status,Shift from Cash_Register where Branch_ID = '" & VBranch_ID & "' And Cash_Register_ID " & _
                " = '" & VReg_ID & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
                If Not RsCek.EOF Then
                    If RsCek!Status <> RsCek!Shift Then
                        MsgBox "X-Reading Belum Di Approve !!!" & vbNewLine & "Harap hubungi Cashier Office.", vbCritical + vbOKOnly, "Oops.."
                        Exit Sub
                    Else
                        If Not Super(3) Then Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            Else
                If Not Super(3) Then Exit Sub
            End If
        End If
        Call XRead
        Call SaveLog("Close Shift " & VSuper_Nm)
        Unload Me
        frmSplash.Show
    Case 2 'CLOSE REGISTER
        If Cfg_Get("Device", "X_ReadPrint", App.Path & "\config.ini") = 1 Then
            If Not Super(2) Then Exit Sub
        Else
            If Linked Then
                RsCek.Open "Select Spending_Program_ID As Status,Shift from Cash_Register where Branch_ID = '" & VBranch_ID & "' And Cash_Register_ID " & _
                " = '" & VReg_ID & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
                If Not RsCek.EOF Then
                    If RsCek!Status <> RsCek!Shift Then
                        MsgBox "Z-Reset Belum Di Approve !!!" & vbNewLine & "Harap hubungi Cashier Office.", vbCritical + vbOKOnly, "Oops.."
                        Exit Sub
                    Else
                        If Not Super(3) Then Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            Else
                If Not Super(3) Then Exit Sub
            End If
        End If
        frmSOD.Caption = "EOD"
        frmSOD.Show 1
        DoEvents
        Call ZReset
        Call SaveLog("Close Register " & VSuper_Nm)
        cmdMenu_Click (9)
    Case 3 'OPEN DRAWER
        'If Cfg_Get("Device", "X_ReadPrint", App.Path & "\config.ini") = 1 Then
            If Not Super(2) Then Exit Sub
        'Else
            'If Not Super(3) Then Exit Sub
        'End If
        Call OpenLaci(1)
        Call SaveLog("Open Drawer " & VSuper_Nm)
    Case 4 'SALES
        VNomor = ""
        Call CDisplay("SALES", "TRANSACTION")
        frmCard.Caption = "SALES"
        frmCard.Show 1
    Case 5 ' REFUND
        VNomor = ""
        If Not Super(1) Then Exit Sub
        Call CDisplay("REFUND", "TRANSACTION")
        Call SaveLog("Refund Transaction " & VSuper_Nm)
        frmCard.Caption = "REFUND"
        frmCard.Show 1
    Case 6 'REPRINT
        If Not Super(1) Then Exit Sub
        Call CDisplay("REPRINT", "TRANSACTION")
        frmNum.Caption = "REPRINT"
        frmNum.Show 1
    Case 7 'RELEASE
        Call CDisplay("RELEASE", "TRANSACTION")
        frmNum.Caption = "RELEASE"
        frmNum.Show 1
    Case 8 'CANCEL
        If Not Super(2) Then Exit Sub
        Call CDisplay("CANCEL", "TRANSACTION")
        frmNum.Caption = "CANCEL"
        frmNum.Show 1
    Case 9 'EXIT
        Call CDisplay("", "")
        frmSOD.Caption = "EOD"
        frmSOD.Show 1
        End
    End Select
End Sub

Private Sub cmdMenu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 78 Then
        frmCekBin.Show 1
        Exit Sub
    End If
    Select Case KeyStroke(KeyCode)
    Case "CASH_OPEN", "CHANGE PASSWORD"
        If cmdMenu(0).Enabled Then cmdMenu_Click (0)
    Case "CLOSE_SHIFT"
        If cmdMenu(1).Enabled Then cmdMenu_Click (1)
    Case "CLOSE_REGISTER"
        If cmdMenu(2).Enabled Then cmdMenu_Click (2)
    Case "OPEN_DRAWER"
        If cmdMenu(3).Enabled Then cmdMenu_Click (3)
    Case "SALES"
        If cmdMenu(4).Enabled Then cmdMenu_Click (4)
    Case "REFUND"
        If cmdMenu(5).Enabled Then cmdMenu_Click (5)
    Case "REPRINT"
        If cmdMenu(6).Enabled Then cmdMenu_Click (6)
    Case "RELEASE"
        If cmdMenu(7).Enabled Then cmdMenu_Click (7)
    Case "CANCEL"
        If cmdMenu(8).Enabled Then cmdMenu_Click (8)
    Case "EXIT"
        If cmdMenu(9).Enabled Then cmdMenu_Click (9)
    End Select
End Sub


Private Sub Form_Activate()
Dim X As Byte

    Star_Nm = ""
    Star_Pt = 0
    Star_Id = ""
    Star_Freq = ""
    Star_Omz = ""
    Call Timer1_Timer

    For X = 1 To 8
        cmdMenu(X).Enabled = VCopen
    Next X

    If Not VCopen Then
        cmdMenu(0).Caption = "CASH OPEN"
        cmdMenu(0).SetFocus
    Else
        cmdMenu(0).Caption = "CHANGE PASSWORD"
        cmdMenu(4).SetFocus
    End If
    
    lblbranch = Tulis(10)
    lblreg = "REGISTER # " & VReg_ID & " - " & VShift
    lblkasir = VKasir_ID & " - " & VKasir_Nm
    lblline = VPing
    txtinfo = Tulis(16)
    
    cmdMenu(9).Enabled = False
    cmdMenu(1).Enabled = False
    cmdMenu(2).Enabled = False
    
    AdoLocal.ConnectionString = StrConLoc

    AdoLocal.RecordSource = "SELECT cast(right(transaction_number,4)as int) as nomor, Net_amount, transaction_number " & _
            "From Sales_Transactions WHERE (Status = '01') and CONVERT(varchar(10), Transaction_Date, 102) = '" & _
            Format(GetSrvDate, "YYYY.MM.DD") & "' and cash_register_id ='" & VReg_ID & "'"
            
    AdoLocal.Refresh
    If AdoLocal.Recordset.EOF Then
        cmdMenu(9).Enabled = True
        cmdMenu(1).Enabled = VCopen
        cmdMenu(2).Enabled = VCopen
    End If

    Call CDisplay("STAR", "DEPARTMENT STORE")
End Sub

Private Sub Form_Load()
    Me.Caption = "POINT OF SALES V." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Grid_Hold_GotFocus()
    cmdMenu(4).SetFocus
End Sub

Private Sub Label1_Click()
Dim Dbs As String, Svr As String

    Dbs = Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini")
    Svr = "[" & VSvr & "]"
    
    If Not Super(2) Then Exit Sub
    ConnLocal.Execute "exec spp_DownLoadOthers '" & Svr & "','" & Dbs & "'"
    Call SaveLog("Download Promo " & VSuper_Nm)
    MsgBox "Download Promo Selesai", vbOKOnly + vbInformation, "Oops.."
End Sub

Private Sub Label2_Click()
    frmCekBin.Show 1
End Sub

Private Sub txtinfo_GotFocus()
    cmdMenu(4).SetFocus
End Sub

Private Sub Timer1_Timer()
    lbltgl = Format(Now, "dddd, d MMM yyyy")
    lbljam = Format(Now, "HH:MM:SS")
End Sub
