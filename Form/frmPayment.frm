VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10875
   ControlBox      =   0   'False
   Icon            =   "frmPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3720
   End
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   2115
      Index           =   3
      Left            =   150
      TabIndex        =   24
      Top             =   3375
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
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   825
         Visible         =   0   'False
         Width           =   840
      End
      Begin TDBNumber6Ctl.TDBNumber txttukar_point 
         Height          =   390
         Left            =   150
         TabIndex        =   26
         Top             =   1350
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":08CA
         Caption         =   "frmPayment.frx":08EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0966
         Keys            =   "frmPayment.frx":0984
         Spin            =   "frmPayment.frx":09CE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0;(#,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,##0;(#,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
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
      Begin TDBNumber6Ctl.TDBNumber txtsaldo_point 
         Height          =   390
         Left            =   3075
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   300
         Width           =   2040
         _Version        =   65536
         _ExtentX        =   3598
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":09F6
         Caption         =   "frmPayment.frx":0A16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0A88
         Keys            =   "frmPayment.frx":0AA6
         Spin            =   "frmPayment.frx":0AF0
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
      Begin TDBText6Ctl.TDBText txtcard_no 
         Height          =   390
         Left            =   150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":0B18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0B84
         Key             =   "frmPayment.frx":0BA2
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
         Left            =   150
         TabIndex        =   25
         Top             =   825
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":0BF4
         Caption         =   "frmPayment.frx":0C14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":0C92
         Keys            =   "frmPayment.frx":0CB0
         Spin            =   "frmPayment.frx":0CFA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,##0;(#,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,##0;(###,##0)"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999
         MinValue        =   -9999999
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
      Begin VB.Label intCicilan 
         Caption         =   "Label3"
         Height          =   135
         Left            =   4800
         TabIndex        =   50
         Top             =   1800
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   4200
      TabIndex        =   44
      Top             =   150
      Width           =   1440
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
         Left            =   150
         Picture         =   "frmPayment.frx":0D22
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1200
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
         Left            =   150
         Picture         =   "frmPayment.frx":1724
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1140
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
         Left            =   150
         Picture         =   "frmPayment.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3240
      Left            =   5550
      TabIndex        =   30
      Top             =   2100
      Width           =   5190
      Begin VB.CommandButton btnNum 
         Caption         =   "ENTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   11
         Left            =   3075
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1125
         Width           =   1950
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   12
         Left            =   4050
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   3075
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   3075
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1950
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   150
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   1125
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   2100
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   150
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1125
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   1125
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1125
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   2100
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1125
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   7
         Left            =   150
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   8
         Left            =   1125
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton btnNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   2100
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   5700
      TabIndex        =   13
      Top             =   150
      Width           =   5040
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
         Left            =   2100
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
         Left            =   2100
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
         X1              =   2400
         X2              =   4875
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
         Left            =   2100
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
         Left            =   75
         TabIndex        =   20
         Top             =   975
         Width           =   4890
      End
   End
   Begin TrueOleDBGrid80.TDBGrid GridPaid 
      Bindings        =   "frmPayment.frx":2B28
      Height          =   2115
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5520
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   3731
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
      Bindings        =   "frmPayment.frx":2B3D
      Height          =   3165
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5583
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
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5318"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5212"
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
      Top             =   6510
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
      Top             =   6870
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
      Top             =   6900
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   556
      Calculator      =   "frmPayment.frx":2B52
      Caption         =   "frmPayment.frx":2B72
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayment.frx":2BDE
      Keys            =   "frmPayment.frx":2BFC
      Spin            =   "frmPayment.frx":2C46
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
      Left            =   150
      TabIndex        =   2
      Top             =   3375
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
         Calculator      =   "frmPayment.frx":2C6E
         Caption         =   "frmPayment.frx":2C8E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2CF2
         Keys            =   "frmPayment.frx":2D10
         Spin            =   "frmPayment.frx":2D5A
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
         ValueVT         =   1996816385
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
         Calculator      =   "frmPayment.frx":2D82
         Caption         =   "frmPayment.frx":2DA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":2E0C
         Keys            =   "frmPayment.frx":2E2A
         Spin            =   "frmPayment.frx":2E74
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
      Left            =   150
      TabIndex        =   7
      Top             =   3375
      Width           =   5265
      Begin VB.CommandButton cmdvoc 
         Height          =   390
         Left            =   3750
         Picture         =   "frmPayment.frx":2E9C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   300
         Width           =   990
      End
      Begin TDBText6Ctl.TDBText txtno_voc 
         Height          =   390
         Left            =   300
         TabIndex        =   8
         Top             =   300
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Caption         =   "frmPayment.frx":389E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":390E
         Key             =   "frmPayment.frx":392C
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
         TabIndex        =   23
         Top             =   840
         Width           =   3390
         _Version        =   65536
         _ExtentX        =   5980
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":3970
         Caption         =   "frmPayment.frx":3990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":39F6
         Keys            =   "frmPayment.frx":3A14
         Spin            =   "frmPayment.frx":3A5E
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
         ValueVT         =   1997996033
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
   End
   Begin VB.Frame frmpay 
      BackColor       =   &H00FFFFFF&
      Height          =   1965
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   3375
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
         Caption         =   "frmPayment.frx":3A86
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":3AF2
         Key             =   "frmPayment.frx":3B10
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
         Caption         =   "frmPayment.frx":3B54
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":3BB8
         Key             =   "frmPayment.frx":3BD6
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
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   688
         Calculator      =   "frmPayment.frx":3C28
         Caption         =   "frmPayment.frx":3C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPayment.frx":3CAE
         Keys            =   "frmPayment.frx":3CCC
         Spin            =   "frmPayment.frx":3D16
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
         ValueVT         =   1983774721
         Value           =   0
         MaxValueVT      =   -1291845627
         MinValueVT      =   1297022981
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   1140
      Left            =   5325
      TabIndex        =   48
      Top             =   2100
      Width           =   465
   End
   Begin VB.Label vstatus 
      Height          =   315
      Left            =   2775
      TabIndex        =   22
      Top             =   6900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label vno_trans 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   6525
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
      Height          =   390
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   3840
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lokasi As String
Dim VTukar_Point As String
Dim PromoInfo, PromoDesc As String
Dim TotPay As Long
Dim RoundOfPay As Long
Dim cnt As Integer
Dim ConCom As New MSComm
Dim Response As String
Dim EcrUse As Boolean


Private Sub cmdvoc_Click()
    frmDisc.lblmsg.Caption = "VOUCHER"
    frmDisc.Show 1
    txtno_voc.SetFocus
    SendKeys "{end}"
End Sub

Private Sub Form_Activate()
    
    TotPay = vpay
    lbltotal = Format(vpay, "#,#00")
    lblbayar = 0
    lblsisa = Format(vpay, "#,#00")
    'Tambahan Harga Point Variable
    txtharga_point.Text = VHargaPoint
    txtno_kartu.Enabled = True
    cnt = 0
    If EcrUse = True Then
        Cek_ECR
    End If
    
End Sub

Private Sub Form_Load()
    EcrUse = False
    Tdata1.ConnectionString = StrConLoc
    Tdata2.ConnectionString = StrConLoc
    vno_trans = VNomor
    If MSCTlp = True Then
        Tdata1.RecordSource = "SELECT Payment_Types, Description, Types, Seq From Payment_Types where Seq<30 And Payment_Types <> 5 ORDER BY Seq"
        Tdata2.RecordSource = "SELECT aa.Seq, bb.Description, Paid_Amount, Credit_Card_No, Credit_Card_Name " & _
                          "FROM Paid aa INNER JOIN Payment_Types bb ON aa.Payment_Types = bb.Payment_Types" & _
                          " where transaction_number = '" & vno_trans & "' order by aa.seq"
    Else
        Tdata1.RecordSource = "SELECT Payment_Types, Description, Types, Seq From Payment_Types where Seq<30 ORDER BY Seq"
        Tdata2.RecordSource = "SELECT aa.Seq, bb.Description, Paid_Amount, Credit_Card_No, Credit_Card_Name " & _
                          "FROM Paid aa INNER JOIN Payment_Types bb ON aa.Payment_Types = bb.Payment_Types" & _
                          " where transaction_number = '" & vno_trans & "' order by aa.seq"
    End If

    
    Tdata1.Refresh
    Tdata2.Refresh
    
    If isECR = 1 Then
        If MsgBox("Gunakan ECR BCA ??", vbYesNo, "Informasi") = vbYes Then
            EcrUse = True
        Else
            EcrUse = False
        End If
    End If
    lokasi = "txtcash"
End Sub

Public Function ZeroPad(ByVal strIn As String, ByVal nLength As Long) As String
   ZeroPad = Right(String(nLength, "0") & strIn, nLength)
End Function

Sub Cek_ECR()
On Error GoTo ErrH
        If isECR = 1 Then
            
'            Label3.Visible = True
            Frame2.Enabled = False
'            cmdpay(0).Enabled = False
             txtno_kartu.Enabled = False
                        txtnama.Enabled = False
                        txtcredit.Enabled = False
            With ConCom
                .CommPort = ECRComm
                .Settings = "9600,o,8,1"
                .InputMode = comInputModeText
            If .PortOpen = False Then
            .PortOpen = True
            
            
'            SerialPort1.Encoding = System.Text.Encoding.GetEncoding(28591)
            Dim inputx As String
            Dim Nilai As String
            Dim NilaiAkhir As String
            Nilai = Int(vpay) & "00"
            NilaiAkhir = ConvertHex(ZeroPad(Nilai, 12))
'            inputx = "0150013031" & ConvertHex(ZeroPad(Nilai, 12)) & "303030303030303030303030" & _
'                "31363838373030363237323031383932202020" & _
'                "32313130" & _
'                "30303030303030302020202020204E4E4E2020202020202020202020202020202020202020202020202020" & _
'                "2020202020202020202020202020202020202020202020202020202020202020202020202020202020202020" & _
'                "2020202020202020202020202003"
                
            'pake kartu
            inputx = "0150013031" & ConvertHex(ZeroPad(Nilai, 12)) & "303030303030303030303030" & _
                "20202020202020202020202020202020202020" & _
                "20202020" & _
                "30303030303030302020202020204E4E4E2020202020202020202020202020202020202020202020202020" & _
                "2020202020202020202020202020202020202020202020202020202020202020202020202020202020202020" & _
                "2020202020202020202020202003"
            inputx = "02" & inputx & LRC(inputx)
            
            Dim cc As String
            Dim I As Integer
            Dim bytes() As Byte
            
            'cara 1 bisa
'            ReDim bytes(0 To Len(inputx))
'            For i = 0 To Len(inputx) - 2 Step 2
'            cc = Asc(Mid(inputx, i + 1, 2))
'                bytes(i) = Asc(Mid(inputx, i + 1, 2))
'                cc = Mid(inputx, i + 1, 2)
'            Next
            
            'cara 2
'            ReDim bytes(0 To Len(inputx))
'            For i = 0 To Len(inputx) - 2 Step 2
'                cc = StrConv(Mid(inputx, i + 1, 2), vbFromUnicode)
'                bytes(i) = StrConv(Mid(inputx, i + 1, 2), vbFromUnicode)
'                cc = Mid(inputx, i, 2)
'            Next
            
             'cara 3
'            cc = StrConv(inputx, vbFromUnicode)
'            bytes() = StrConv(inputx, vbFromUnicode)

             'cara 4 oke
            Dim counter As Integer
            ReDim bytes(0 To (Len(inputx) / 2) - 1) As Byte
            For I = 1 To Len(inputx) Step 2
                bytes(counter) = CDbl(Val("&H" & Mid(inputx, I, 2)))
                counter = counter + 1
            Next I
        
               
            .Output = bytes()
            .PortOpen = False
            Timer1.Enabled = True
            GridPay_Types.MoveNext
            Call Timer1_Timer
            End If
            End With
        Else
            Timer1.Enabled = False
        End If
        Exit Sub
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Timer1.Enabled = False
    Frame2.Enabled = True
cmdpay(0).Enabled = True
    End Sub
    
Private Function ConvertHex(ByVal Data As String) As String
        Dim Hasil As String
        Hasil = ""
        Dim x As Integer
        For x = 0 To Len(Data) - 1
            Hasil = Hasil & "3" & Mid(Data, x + 1, 1)
        Next
        Data = Hasil
        ConvertHex = Hasil
End Function

Private Function LRC(ByVal Data As String) As String
        Dim binStr As String
        Dim bin(8) As Integer
        Dim Decimals As String
        Dim Hexadecimal As String
        Dim ByteArray() As Byte
        Dim x As Integer
        For x = 0 To Len(Data) - 2 Step 2
'            binStr = ZeroPad(String(Int(Mid(Data, x + 1, 2)), 2), 8)
            Dim Number As String
            Dim Binary_String As String
            Number = Mid(Data, x + 1, 2)
            
            Binary_String = ""
            Binary_String = HexToBin(Mid(Data, x + 1, 2))
'            While Number > 0
'                Binary_String = Str(Number Mod 2) & Binary_String
'                Number = Number \ 2
'            Wend
            binStr = ZeroPad(Replace(Binary_String, Space(1), ""), 8)
            Dim I As Integer
            For I = 0 To 7
                bin(I + 1) = bin(I + 1) + Mid(binStr, I + 1, 1)
            Next
        Next
        Decimals = bin(1) Mod 2 & bin(2) Mod 2 & bin(3) Mod 2 & bin(4) Mod 2 & bin(5) Mod 2 & bin(6) Mod 2 & bin(7) Mod 2 & bin(8) Mod 2
        Hexadecimal = BinToHex(Decimals)
        LRC = Hexadecimal
    End Function
    
    Public Function HexToString(ByVal HexToStr As String) As String
        Dim strTemp   As String
        Dim strReturn As String
        Dim I         As Long
        For I = 1 To Len(HexToStr) Step 2
            strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
            strReturn = strReturn & strTemp
        Next I
        HexToString = strReturn
    End Function
    
    Public Function StringToHex(ByVal StrToHex As String) As String
        Dim strTemp   As String
        Dim strReturn As String
        Dim I         As Long
        For I = 1 To Len(StrToHex)
            strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
            If Len(strTemp) = 1 Then strTemp = "0" & strTemp
            strReturn = strReturn & Space$(1) & strTemp
        Next I
        StringToHex = strReturn
    End Function
    
Public Function HexToBin(ByVal HexStr As String) As Double
    Dim DecNum As String
    Dim char As String
    Dim I As Integer
    I = 1
    DecNum = ""
    For I = 1 To Len(HexStr)
        char = Mid(HexStr, I, 1)
       If char = "0" Then char = "0000"
If char = "1" Then char = "0001"
If char = "2" Then char = "0010"
If char = "3" Then char = "0011"
If char = "4" Then char = "0100"
If char = "5" Then char = "0101"
If char = "6" Then char = "0110"
If char = "7" Then char = "0111"
If char = "8" Then char = "1000"
If char = "9" Then char = "1001"
If char = "A" Then char = "1010"
If char = "B" Then char = "1011"
If char = "C" Then char = "1100"
If char = "D" Then char = "1101"
If char = "E" Then char = "1110"
If char = "F" Then char = "1111"
        DecNum = DecNum & char
    Next I
    HexToBin = DecNum
    End Function
    
    Public Function BinToHex(ByVal HexStr As String) As String
    Dim DecNum As String
    Dim char As String
    Dim I As Integer
    I = 1
    DecNum = ""
    For I = 1 To Len(HexStr) Step 4
        char = Mid(HexStr, I, 4)
       If char = "0000" Then char = "0"
If char = "0001" Then char = "1"
If char = "0010" Then char = "2"
If char = "0011" Then char = "3"
If char = "0100" Then char = "4"
If char = "0101" Then char = "5"
If char = "0110" Then char = "6"
If char = "0111" Then char = "7"
If char = "1000" Then char = "8"
If char = "1001" Then char = "9"
If char = "1010" Then char = "A"
If char = "1011" Then char = "B"
If char = "1100" Then char = "C"
If char = "1101" Then char = "D"
If char = "1110" Then char = "E"
If char = "1111" Then char = "F"
        DecNum = DecNum & char
    Next I
    BinToHex = DecNum
    End Function


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
    RoundOfPay = 0
    lbltotal = Format(TotPay, "#,#00")
    lblsisa = Format(vpay, "#,#00")
        Select Case GridPay_Types.Columns(2)
            Case "CS"
                txtcash = vpay
                
                RoundOfPay = vpay - (Int(vpay / 100)) * 100
                If vpay < 0 Then
                    If vpay <> Int(vpay / 100) * 100 Then
                        RoundOfPay = vpay - (Int(vpay / 100) + 1) * 100
                    Else
                        RoundOfPay = 0
                    End If
                End If
                'RoundOfPay = 0
                txtcash = Format(vpay - RoundOfPay, "#,#00")
                lbltotal = Format(TotPay - RoundOfPay, "#,#00")
                lblsisa = Format(vpay - RoundOfPay, "#,#00")
                txtcash.SetFocus
            Case "CC", "DC"
                txtcredit = vpay
                txtno_kartu = ""
                txtnama = ""
                txtno_kartu.SetFocus
            Case "SV", "OV"
                txtno_voc = ""
                txtvoucher = 0
                txtno_voc.SetFocus
            Case "PR"
                txtpoint.SetFocus
                txtpoint = Format(txttukar_point * txtharga_point, "#,###,##0")
        End Select
        DoEvents
        SendKeys "{home}+{end}"
    Case 27
        Dim RsHapus As New ADODB.Recordset, RsAmbil As New ADODB.Recordset
        Dim Sisa_Bonus As Long, Bonus As Long
        If Not Tdata2.Recordset.EOF Then Tdata2.Recordset.MoveFirst
        While Not Tdata2.Recordset.EOF
            If Tdata2.Recordset!Description = "POINT REWARD" Then
            
                ConnLocal.Execute "Update card set card_point=card_point + " & Cari_Point(Tdata2.Recordset!credit_card_name) & "where card_nr = '" & Tdata2.Recordset!Credit_Card_No & "'"
                ConnServer.Execute "Update card set card_point=card_point + " & Cari_Point(Tdata2.Recordset!credit_card_name) & "where card_nr = '" & Tdata2.Recordset!Credit_Card_No & "'"
                
                ConnLocal.Execute "delete from cust_point_trans where trans_nr='" & Tdata2.Recordset!credit_card_name & "'"
                ConnServer.Execute "delete from cust_point_trans where trans_nr='" & Tdata2.Recordset!credit_card_name & "'"
            ElseIf Tdata2.Recordset!Description = "DISCOUNT BY STAR" Then
                    If DiscStarProID = 3 Then
                    RsAmbil.Open "Select b.ext1 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
                    Bonus = RsAmbil!ext1
                    If Left(txtcard_no, 5) = "CM999" Then
                        Sisa_Bonus = Bonus + Tdata2.Recordset!paid_amount
                        ConnServer.Execute "Update b set b.ext1 = " & Sisa_Bonus & " From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                    End If
                    End If
                    If DiscStarProID = 9 Then
                        ConnServer.Execute "Update b set b.ext1 = 1 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                        ConnLocal.Execute "Update b set b.ext1 = 1 From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & txtcard_no & "'"
                    End If
            End If
            Tdata2.Recordset.MoveNext
        Wend
        Call MySTAR(txtcard_no, 0)
        Call tampil_point
        
        Call SQLQuery("delete from Paid where Transaction_Number = '" & vno_trans & "'")
    
    VNomor = vno_trans
    Unload Me
    End Select
End Sub

Private Function Cari_Point(No_trans As String) As Integer
Dim RsCari As New ADODB.Recordset
    
    RsCari.Open "select claim_point from cust_point_trans where Trans_nr = '" & No_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
    Cari_Point = IIf(Not IsNull(RsCari!claim_point), RsCari!claim_point, 0)
    RsCari.Close:   Set RsCari = Nothing

End Function

Private Sub GridPay_Types_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim x As Byte
    lblmsg.Caption = GridPay_Types.Columns(1)
    
    For x = 0 To 3
        frmpay(x).Visible = False
    Next x
    RoundOfPay = 0
    lbltotal = Format(TotPay, "#,#00")
    lblsisa = Format(vpay, "#,#00")
    txtkembali = 0
    Select Case GridPay_Types.Columns(2)
        Case "CS"
            frmpay(0).Visible = True
            RoundOfPay = vpay - (Int(vpay / 100)) * 100
            If vpay < 0 Then
                If vpay <> Int(vpay / 100) * 100 Then
                    RoundOfPay = vpay - (Int(vpay / 100) + 1) * 100
                Else
                    RoundOfPay = 0
                End If
            End If
            'RoundOfPay = 0
            lbltotal = Format(TotPay - RoundOfPay, "#,#00")
            lblsisa = Format(vpay - RoundOfPay, "#,#00")
            txtcash = Format(vpay - RoundOfPay, "#,#00")
        Case "CC", "DC"
            frmpay(1).Visible = True
        Case "SV", "OV"
            frmpay(2).Visible = True
        Case "PR"
            Call tampil_point
            frmpay(3).Visible = True
    End Select
    If TipKom = 1 Then Call GridPay_Types_KeyDown(13, 0)
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrH
        cnt = cnt + 1
        If cnt > 2000 Then
            Timer1.Enabled = False
'            Label3.Visible = False
            If MsgBox("Connections Time Out !!! Try Again??", vbYesNo) = vbYes Then
                cnt = 11
'                Label3.Visible = True
                Timer1.Enabled = True
            Else
                Timer1.Enabled = False
                Timer1.Enabled = False
                Frame2.Enabled = True
'                Label3.Visible = False
'                cmdpay(0).Enabled = True
txtno_kartu.Enabled = True
txtnama.Enabled = True
txtcredit.Enabled = True
                Exit Sub
            End If
        End If

        If cnt > 10 Then
            If ConCom.PortOpen = False Then ConCom.PortOpen = True
            Dim buff As String
            Dim I As Integer
            Dim d As String
       

            Response = ""
            buff = ConCom.Input

            For I = 1 To Len(buff)
                d = Hex(Asc(Mid(buff, I, 1)))
                Response = Response & d
            Next I

            If Response = "6" Then
                Response = ""
            End If

            If Response <> "" Then
                    ConCom.PortOpen = False
                    Dim ResponseASCII As String
                    ResponseASCII = HexToString(Mid(Response, 9, Len(Response) - 12))
                    If Mid(ResponseASCII, 48, 2) <> "00" Then
                        Timer1.Enabled = False
                        Timer1.Enabled = False
                        MsgBox ("Respon Failed From EDC Device !!!")
                        Frame2.Enabled = True
'                        Label3.Visible = False
'                        cmdpay(0).Enabled = True
                        txtno_kartu.Enabled = True
                        txtnama.Enabled = True
                        txtcredit.Enabled = True
                        Exit Sub
                    End If
                    Timer1.Enabled = False
                    txtno_kartu = Left(Mid(ResponseASCII, 25, 19), 16)
                    If Mid(ResponseASCII, 163, 1) = "Y" Then
                        txtnama.Text = "*" & Mid(ResponseASCII, 178, 3) & "*"
                    End If
                    txtno_kartu.Enabled = False
                    txtcredit = vpay
                    Dim x As Byte
                     For x = 0 To 3
                        frmpay(x).Visible = False
                    Next x
                frmpay(1).Visible = True
                
                Frame2.Enabled = True
'               Label3.Visible = False
                isECR = 2
'               txtno_kartu.Enabled = True
'               txtnama.Enabled = True
                txtcredit.Enabled = True
                txtcredit.SetFocus
'               txtcredit.Enabled = False
'lokasi = "txtno_kartu"
btnNum_Click (11)
                End If
            End If
            Exit Sub
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Timer1.Enabled = False
    Frame2.Enabled = True
'    cmdpay(0).Enabled = True
     txtno_kartu.Enabled = True
                        txtnama.Enabled = True
                        txtcredit.Enabled = True
End Sub

Private Sub txtcash_Change()
    txtkembali = Val(txtcash + RoundOfPay) - vpay
End Sub

Private Sub txtcash_GotFocus()
    lokasi = "txtcash"
End Sub

Private Sub txtcash_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            GridPay_Types.SetFocus
            If vstatus = "SALES" And txtcash >= 0 Then
                If Simpan_Detail(IIf(txtkembali >= 0, txtcash - txtkembali, txtcash), "", "") Then
                    Call Cek_Lunas
                End If
            ElseIf vstatus = "REFUND" And txtcash <= 0 Then
                If Simpan_Detail(IIf(txtkembali <= 0, txtcash - txtkembali, txtcash), "", "") Then
                    Call Cek_Lunas
                End If
            End If
        Case 27
            GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtcash_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then KeyAscii = 0    '45 -
End Sub

Private Sub txtcredit_GotFocus()
    lokasi = "txtcredit"
End Sub

Private Sub txtno_kartu_GotFocus()
    lokasi = "txtno_kartu"
End Sub

Private Sub txtno_kartu_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Select Case GridPay_Types.Columns(2)
        Case "CC", "DC"
            If txtno_kartu.Text = "" Or Mid(txtno_kartu.Text, 3, 1) = "-" Or Len(txtno_kartu) <> 16 Then
                MsgBox "Nomor kartu tidak valid", vbOKOnly + vbInformation, "Oops.."
                txtno_kartu.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
            Dim Kartu As String

            Kartu = LTrim(Left(txtno_kartu.Text, 79))

'            If Mid(Kartu, 4, 1) = "B" Then
'                txtno_kartu = LTrim(Mid(Kartu, 5, 16))
'                txtnama = LTrim(Mid(Kartu, 22, 20))
'            ElseIf Mid(Kartu, 2, 1) = "B" Then
'                txtno_kartu = LTrim(Mid(Kartu, 3, 16))
'            ElseIf Left(Kartu, 1) = ";" Then
'                txtno_kartu = LTrim(Mid(Kartu, 2, 16))
'            ElseIf Mid(Kartu, 4, 1) <> "B" Then
'                txtno_kartu = LTrim(Mid(Kartu, 8, 16))
'            End If
            '---------------------- normal ----------------
            If Mid(Kartu, 4, 1) <> "B" Then
                If Len(Kartu) = 16 Then
                    txtno_kartu = Kartu
                Else
                    If Left(Kartu, 1) <> "c" Then Call SaveLog(Kartu & " - KARTU BARU")
                    If Mid(Kartu, 8, 1) = "B" Then
                        txtno_kartu = LTrim(Mid(Kartu, 9, 16))
                    Else
                        txtno_kartu = LTrim(Mid(Kartu, 8, 16))
                    End If
               End If
            ElseIf Mid(Kartu, 4, 1) = "B" Then
                txtno_kartu = LTrim(Mid(Kartu, 5, 16))
                txtnama = LTrim(Mid(Kartu, 22, 20))
           End If
                        '---------------POSIFLEX------------------------
'           If Left(Kartu, 1) = ";" Then
'                txtno_kartu = LTrim(Mid(Kartu, 2, 16))
'           ElseIf Mid(Kartu, 2, 1) <> "B" Then
'                'Call SaveLog(Kartu & " - KARTU BARU")
'                txtno_kartu = LTrim(Mid(Kartu, 8, 16))
'           ElseIf Mid(Kartu, 2, 1) = "B" Then
'                txtno_kartu = LTrim(Mid(Kartu, 3, 16))
'                txtnama = LTrim(Mid(Kartu, 20, 26))
'           End If

'       ---------------NEC------------------------
'           If Mid(Kartu, 2, 1) = "B" Then
'                'Call SaveLog(Kartu & " - KARTU BARU")
'                txtno_kartu = LTrim(Mid(Kartu, 3, 16))
'                txtnama = LTrim(Mid(Kartu, 20, 26))
'           ElseIf Left(Kartu, 1) = ";" Then
'                txtno_kartu = LTrim(Mid(Kartu, 2, 16))
'           ElseIf Mid(Kartu, 2, 1) <> "B" Then
'                txtno_kartu = LTrim(Mid(Kartu, 8, 16))
'           End If

            If Len(txtno_kartu) <> 16 Then Call SaveLog(Kartu & " - TIDAK16")
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
            If (vstatus = "SALES" And txtcredit <= vpay) Or (vstatus = "REFUND" And txtcredit >= vpay) Then
                GridPay_Types.SetFocus
                'Cek_Promo_Card
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

Private Sub txtno_voc_GotFocus()
    lokasi = "txtno_voc"
End Sub
Private Sub txtno_voc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RsDobel As New ADODB.Recordset

    Select Case KeyCode
    Case 13
        Select Case GridPay_Types.Columns(2)
        Case "SV"
                     
            If txtno_voc = "50" Then
                txtvoucher = "50000"
                txtvoucher.SetFocus
                Exit Sub
            End If
            
            If txtno_voc = "75" Then
                txtvoucher = "75000"
                txtvoucher.SetFocus
                Exit Sub
            End If
            
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
                txtno_voc.Text = ""
                txtno_voc.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            End If
            RsDobel.Close:   Set RsDobel = Nothing
            
            If Len(VKary) = 9 Then
                MsgBox "Promo disc karyawan tidak bisa dibayar dengan voucher", vbOKOnly + vbInformation, "Oops.."
                txtno_voc.Text = ""
                txtno_voc.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            End If
            
'--VOUCHER youngSTYLE tidak berlaku saat Late Nite
'            If Left(txtno_voc, 2) = "ZB" Then
'            Dim RsZB As New ADODB.Recordset
'                StrSQL = "SELECT * from promo_hdr WHERE getdate() Between Start_Date And End_Date and aktif=1 and tipe=30"
'                If Linked Then
'                    RsZB.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
'                Else
'                    RsZB.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
'                End If
'
'                If Not RsZB.EOF Then
'                    MsgBox "Voucher youngSTYLE tidak berlaku saat LATE NITE", vbOKOnly + vbInformation, "Oops.."
'                    txtno_voc.SetFocus
'                    DoEvents
'                    SendKeys "{home}+{end}"
'                End If
'                RsZB.Close: Set RsZB = Nothing
'            End If
'--VOUCHER youngSTYLE tidak berlaku saat Late Nite

'--VOUCHER FKB tidak boleh gabung voucher lain
            If Left(txtno_voc, 2) = "ZR" Then StrSQL = "select * from paid where transaction_number='" & vno_trans & "' and payment_types=8 and LEFT(Credit_Card_No, 2) <> 'ZR'"
            If Left(txtno_voc, 2) <> "ZR" Then StrSQL = "select * from paid where transaction_number='" & vno_trans & "' and payment_types=8 and LEFT(Credit_Card_No, 2) = 'ZR'"
            
            Dim RsZB As New ADODB.Recordset
                RsZB.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly

                If Not RsZB.EOF Then
                    MsgBox "Voucher FKB tidak bisa digabung dengan Voucher STAR Lainnya", vbOKOnly + vbInformation, "Oops.."
                    txtno_voc.SetFocus
                    DoEvents
                    SendKeys "{home}+{end}"
                End If
                RsZB.Close: Set RsZB = Nothing
'--VOUCHER FKS dan Jumbo Cash Back tidak bisa bareng
        Case "OV"
            txtvoucher.ReadOnly = False
            txtvoucher.SetFocus
            DoEvents
            SendKeys "{home}+{end}"
        End Select
    Case 27
        GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txtvoucher_GotFocus()
    lokasi = "txtvoucher"
End Sub

Private Sub txtvoucher_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        If (vstatus = "SALES" And txtvoucher > 0) Or (vstatus = "REFUND" And txtvoucher < 0) Then
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
        Case "OV"
        If GridPay_Types.Columns(1) = "GO PAY" Then
            If Int(txtvoucher.Value) < 150000 Then
                MsgBox "Pembayaran menggunakan 'GO PAY' minimal Rp. 150.000 !!", vbOKOnly + vbInformation, "Oops.."
                txtvoucher.ReadOnly = False
                txtvoucher.SetFocus
                Exit Sub
            End If
        End If
        
        If GridPay_Types.Columns(1) = "DANA" Then
            If Int(txtvoucher.Value) < 150000 Then
                MsgBox "Pembayaran menggunakan 'DANA' minimal Rp. 150.000 !!", vbOKOnly + vbInformation, "Oops.."
                txtvoucher.ReadOnly = False
                txtvoucher.SetFocus
                Exit Sub
            End If
        End If
        
        If GridPay_Types.Columns(1) = "OVO" Then
            If Int(txtvoucher.Value) < 150000 Then
                MsgBox "Pembayaran menggunakan 'OVO' minimal Rp. 150.000 !!", vbOKOnly + vbInformation, "Oops.."
                txtvoucher.ReadOnly = False
                txtvoucher.SetFocus
                Exit Sub
            End If
        End If
        End Select
        
        GridPay_Types.SetFocus
            'If txtvoucher = "25000" Or txtvoucher = "50000" Or txtvoucher = "100000" Then
            If Simpan_Detail(txtvoucher, txtno_voc, GridPay_Types.Columns(1)) Then
                txtno_voc = ""
                txtvoucher = 0
                Call Cek_Lunas
            End If
            'Else
            '    MsgBox "Nilai voucher tidak valid", vbOKOnly + vbInformation, "Oops.."
            '    txtvoucher.SetFocus
            '    SendKeys "{home}+{end}"
            'End If
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
        txtno_voc.Text = ""
        Exit Function
    End If
    RsCari.Close:   Set RsCari = Nothing

    RsDobel.Open "select credit_card_no from paid where credit_card_no = '" & nomor & "'", ConnServer, adOpenForwardOnly, adLockReadOnly
    If Not RsDobel.EOF Then
        MsgBox "Nomor Voucher Sudah Pernah Dipakai", vbOKOnly + vbInformation, "Oops.."
        Cek_Voc = 0
        txtno_voc.Text = ""
    End If
    RsDobel.Close:   Set RsDobel = Nothing
End Function

Private Function Simpan_Detail(bayar As Double, card_no As String, card_num As String) As Boolean
On Error GoTo ErrH
    
    Simpan_Detail = False
    
    intCicilan = 1
    If InStr(GridPay_Types.Columns(1), "CICILAN") > 0 Then
        frmCicilan.Show 1
        DoEvents
    End If
    
    ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No, " & _
            "Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & vno_trans & "','" & GridPay_Types.Columns(0) & "','" & Gen_Seq & _
            "','IDR','" & intCicilan & "','" & card_no & "','" & UbahChar(card_num) & "'," & bayar & ",'" & VShift & "')"
    
    Tdata2.RecordSource = "SELECT aa.Seq, bb.Description, Paid_Amount, Credit_Card_No, Credit_Card_Name " & _
                        "FROM Paid aa INNER JOIN Payment_Types bb ON aa.Payment_Types = bb.Payment_Types" & _
                        " where transaction_number = '" & vno_trans & "' order by seq"
    Tdata2.Refresh
    GridPaid.MoveLast
    
    'vpay = vpay - (bayar + TotPay)
    If RoundOfPay <> 0 And (bayar + lblbayar + RoundOfPay) < TotPay Then
    vpay = vpay - (bayar)
    lblbayar = Format(lblbayar + bayar, "#,##0")
    lblsisa = Format(lbltotal - lblbayar, "#,##0")
    Else
    vpay = vpay - (bayar + RoundOfPay)
    lblbayar = Format(lbltotal - vpay, "#,##0")
    lblsisa = Format(vpay, "#,##0")
    End If
    
    
    'TotPay = 0
    Simpan_Detail = True
    Exit Function
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Simpan_Detail " & Err.Description & " " & Err.Number)
End Function

'Private Sub Cek_Promo_Card()
'Dim RsPromo As New ADODB.Recordset
'Dim NilaiOK As String
'Dim Cek_KK As New ADODB.Recordset

'StrSQL = "SELECT promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn " & _
'             ", SUM(Sales_Transaction_Details.Net_Price) As Belanja, islimit, qtylimit, qtyout, " & _
'             " isnull(txt1,'') as txt1, isnull(txt2,'') as txt2, isnull(txt3,'') as txt3, isnull(txt4,'') as txt4 FROM Promo_Hdr " & _
'             "INNER JOIN Promo_Dtl ON Promo_Hdr.promo_id = Promo_Dtl.promo_id " & _
'             "INNER JOIN Sales_Transaction_Details ON Promo_Dtl.PLU = Sales_Transaction_Details.PLU " & _
'             "WHERE (Sales_Transaction_Details.Transaction_Number = '" & vno_trans & "') " & _
'             "AND getdate() Between Start_Date And End_Date and aktif=1 " & _
'             "GROUP BY promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn, " & _
'             "islimit, qtylimit, qtyout, txt1, txt2, txt3, txt4 Having (promo_hdr.tipe = 49) and SUM(Sales_Transaction_Details.Net_Price)>0 " & _
'             "order by promo_hdr.promo_id"
    
'    RsPromo.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
'    If RsPromo.EOF Then
'        RsPromo.Close: Set RsPromo = Nothing
'        Exit Sub
'    End If
    
    
'    StrSQL = "select CAST(nomor AS varchar(6)) from cc_master where cc_master='" & RsPromo!promo_id & "' and CAST(nomor AS varchar(6)) = '" & Left(txtno_kartu.Text, 6) & "'"
'    Cek_KK.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    
'    If Cek_KK.EOF Then
'    Cek_KK.Close
'       Exit Sub
'    End If
'    Cek_KK.Close
    
'    NilaiOK = Pakai_KK(RsPromo!promo_id, vno_trans)
    
'    If NilaiOK > 0 Then
'        MsgBox "Kartu ini mendapat diskon  " & RsPromo!disc & _
'                "% ", vbOKOnly + vbInformation, "Oops.."
'                TotPay = (RsPromo!disc / 100) * txtcredit.Text
'        txtcredit.Text = FormatNumber(txtcredit.Text - (RsPromo!disc / 100) * txtcredit.Text, 0)
'        ConnLocal.Execute "INSERT Into Paid(Transaction_Number, Payment_Types, Seq, Currency_ID, " & _
'                    "Currency_Rate, Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) VALUES ('" & _
'                    vno_trans & "','31','" & Gen_Seq & "','IDR','1','','" & RsPromo!promo_name & "'," & TotPay & ",'" & VShift & "')"

'    End If
'End Sub

Private Sub Cek_Lunas()
On Error GoTo ErrH
Dim q As Byte, k As Long
Dim l As Long
Dim QtyOutStr As Long
Dim RsInfoPromoSeq As New ADODB.Recordset
Dim RsCari As New ADODB.Recordset
    If vpay <= 0 And vstatus = "SALES" Then 'jika lunas
        If VCekKartu <> "" Then
            If Pakai_KK(VCekKartu, vno_trans) <= 0 Then
                MsgBox "Transaksi ini harus menggunakan kartu kredit/debit yang sedang promo", vbCritical + vbOKOnly, "Oops.."
                Exit Sub
            Else
                If isLimitCC = 1 Then
                QtyOutStr = (CDec(VDiscBySTAR) / 1000)
                ConnLocal.Execute "Update Promo_hdr set QtyOut = QtyOut + " & QtyOutStr & " where promo_id = '" & VCekKartu & "'"
                If Linked Then
                    ConnServer.Execute "Update Promo_hdr set QtyOut = QtyOut + " & QtyOutStr & " where promo_id = '" & VCekKartu & "'"
                End If
                isLimitCC = 0
                End If
            End If
        End If
        If RoundOfPay <> 0 Then
        ConnLocal.Execute "Insert Into Paid select top 1 '" & vno_trans & "','36',(Select top 1 seq + 1 from paid where transaction_number = '" & vno_trans & "' Order By Seq desc),'IDR','1','','ROUNDING'," & RoundOfPay & ",Shift from paid where transaction_number = '" & vno_trans & "'"
        End If
        
        Call Paid_To_Sales(vno_trans, lbltotal + RoundOfPay, IIf(txtkembali > 0, txtkembali, 0))
        GoTo lanjut
    End If
    
    If vpay >= 0 And vstatus = "REFUND" Then 'jika lunas
        If RoundOfPay <> 0 Then
        ConnLocal.Execute "Insert Into Paid select top 1 '" & vno_trans & "','36',(Select top 1 seq + 1 from paid where transaction_number = '" & vno_trans & "' Order By Seq desc),'IDR','1','','ROUNDING'," & RoundOfPay & ",Shift from paid where transaction_number = '" & vno_trans & "'"
        End If
        Call Paid_To_Sales(vno_trans, lbltotal, IIf(txtkembali < 0, txtkembali, 0))
        GoTo lanjut
    End If
    txtcash = Format(vpay - RoundOfPay, "#,#00")
    txtkembali = 0
    Exit Sub

lanjut:
    'tambahan update v_flag voucher automatic
    If Linked Then
    Dim Dbs As String, Svr As String
    Svr = "[" & VSvr & "]"
    Dbs = Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini")
    ConnLocal.Execute "exec spp_Upload_to_Server " & Svr & ", " & Dbs & ", '" & vno_trans & "'"
'    Call Upload_to_Server(vno_trans)
'    ConnServer.Execute "Update newvoc set v_flag = 'R',V_REC = CONVERT(VARCHAR(10), GETDATE(), 120)  where V_NO in (Select credit_card_no from paid where payment_types = 8 and transaction_number = '" & vno_trans & "')"
    End If
    If txtcard_no <> "CM000-00000" Then Call Save_Point(vno_trans, txtcard_no)
    Call OpenLaci(0) ' buka drawer tanpa print
    
    'Penambahan Text Acara Marketing
    If AcaraMKT = 1 Then
        RsCari.Open "select Sum(Qty) as Total from Sales_Transaction_Details where Transaction_Number = '" & _
                    vno_trans & "' And PLU in ('9000092900005')", ConnLocal, adOpenForwardOnly, adLockReadOnly
        Footer6str = IIf(IsNull(RsCari!total), 500, RsCari!total * 10000 + 500)
        RsCari.Close:   Set RsCari = Nothing
    End If
    
    If PathEmail <> "" Then
        If StrukEmail = True Then
            Dim intFile As Integer
            Dim strFile As String
            strFile = PathEmail & "\JgnDihapus.txt"
            intFile = FreeFile
            Open strFile For Output As #intFile
            Print #intFile, "Oke"
            Close #intFile
            Kill PathEmail & "\*.*"
ulang:
            'InputEmail = InputBox("Masukan Email ?", "Email", Star_Email)
            
            Call CetakStrukEmail(vstatus, vno_trans)
        Else
            Call CetakStruk(vstatus, vno_trans)
        End If
    Else
        Call CetakStruk(vstatus, vno_trans)
    End If
    

      
    'If Linked Then Call CetakStrukPayPoint(txtcard_no, Left(Star_Nm, 22), vno_trans)
    
    If Mid(txtcard_no, 1, 5) = "CM999" Then
    If Linked Then
        l = Cek_BonusKaryawan(txtcard_no, vno_trans)
    Else
        l = 0
    End If
    Else
    l = 0
    End If
    
    If l = 0 Then
    If vstatus = "SALES" Then
    Call CetakPromo(vno_trans, 0)
    If StrukEmail = True Then
        Me.Enabled = False
        '-----kirim email
        ''Call SendEmail(StoreEmail, InputEmail, vno_trans)
        '-----Copy file
        If Not Dir(EReceiptEmail & vno_trans & " .pdf") = "" Then
            Kill EReceiptEmail & vno_trans & " .pdf"
        End If
        'FileCopy PathEmail & "\" & vno_trans & ".pdf", EReceiptEmail & vno_trans & ".pdf"
        '-----Copy file
        Dim Fso As New FileSystemObject
        Dim fil As File

        For Each fil In Fso.GetFolder(PathEmail).Files
            Debug.Print
            Dim TabFile() As String
            TabFile = Split(fil.name, ".")
            If TabFile(1) = "jpg" Or TabFile(1) = "bmp" Then
                Kill PathEmail & "\" & fil.name
            Else
                'If Not Dir(PathEmail & "\Backup\" & fil.name) = "" Then
                '    Kill PathEmail & "\Backup\" & fil.name
                'End If
                If Not Dir(EReceiptEmail & fil.name) = "" Then
                    Kill EReceiptEmail & fil.name
                End If
                FileCopy PathEmail & "\" & fil.name, EReceiptEmail & fil.name
                FileCopy PathEmail & "\" & fil.name, PathEmail & "\BACKUP\" & fil.name
            End If
            
        Next
        ConnServer.Execute "Insert Into EReceipt_Email (Email,Trans_Nr,Gender,Nama,Tanggal,Reprint,status) values ('" & InputEmail & "','" & vno_trans & "','" & Star_Gender & "', '" & InputNama & "',getdate(),0, 0)"
        
        Dim RsCari2 As New ADODB.Recordset
        StrSQL = "select * from EReceipt_Email_Contact where Email = '" & InputEmail & "'"
        If Linked Then
            RsCari2.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
        Else
            RsCari2.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
        End If
        If RsCari2.EOF Then
           ConnServer.Execute "Insert Into EReceipt_Email_Contact (Email,Nama,Hp) values ('" & InputEmail & "','" & InputNama & "', '" & InputNoTlp & "')"
        End If
        Me.Enabled = True
    End If
    
    If VPing = "ONLINE" Then
        If Star_No <> "CM000-00000" Then
            RsInfoPromoSeq.Open "select a.promo_id from Promo_Hdr a inner join Seqmentation_Member_Promo b on a.promo_id = b.promo_id where b.card_nr " & _
            " = '" & Star_No & "' and getdate() Between a.Start_Date And a.End_Date and a.aktif=1 and b.status = 1 and a.tipe >= 31 and a.seqmentation <> 0", ConnServer, adOpenForwardOnly, adLockReadOnly
            While Not RsInfoPromoSeq.EOF
                UpdateStatusSeq = True
                Call CetakPromo_Seq(vno_trans, 1, RsInfoPromoSeq!promo_id)
                If UpdateStatusSeq = True Then
                    ConnServer.Execute "Update Seqmentation_Member_Promo set Status=0 where promo_id = '" & RsInfoPromoSeq!promo_id & "' and Card_Nr='" & Star_No & "'"
                End If
                RsInfoPromoSeq.MoveNext
            Wend
            RsInfoPromoSeq.Close: Set RsInfoPromoSeq = Nothing
        End If
    End If
    End If
    End If
    Dim ix As Integer
    If SeqCountInt > 0 Then
    For ix = 1 To SeqCountInt
        ConnServer.Execute "Update Seqmentation_Member_Promo set Status=0 where promo_id = '" & SeqCount(ix) & "' and Card_Nr='" & Star_No & "'"
    Next
    SeqCountInt = 0
    End If
    
    'If vstatus = "SALES" Then Call UpdatePromo(vno_trans)
    

     
    If Linked Then Call CetakStrukPayPoint(txtcard_no, Left(Star_Nm, 22), vno_trans)
    If Linked Then Call CetakStrukEVoucher(vno_trans)
    Call CetakStrukDANA(vno_trans)
    
    For q = 0 To 4
        frmSales.cmdsales(q).Enabled = False
    Next q
    
    frmSales.cmdsales(8).Enabled = False
    frmSales.cmdsales(9).Enabled = False
    
    frmSales.CmdNav(0).Enabled = False
    frmSales.CmdNav(3).Enabled = False
    
    
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
    'VNomor = ""
    Exit Sub
        
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Cek Lunas " & Err.Description & " " & Err.Number)
End Sub

Private Sub UpdatePromo(No_trans As String)
'On Error Resume Next
'Dim RsUpdPro As New ADODB.Recordset
'    RsUpdPro.Open "select * from Sales_Transaction_Details sd inner join item_master im on sd.plu=im.plu " & _
'                  "where im.Burui='NMD92ZZZ9' and Discount_Percentage>0 and Transaction_Number='" & No_trans & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
'
'    If Not RsUpdPro.EOF Then
'        StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('MV1', '" & Star_No & "', 0, 0, '00')"
'
'        ConnLocal.Execute StrSQL
'        If Linked Then
'            ConnServer.Execute StrSQL
'            ConnLocal.Execute "Update promo_sales set status='99' where promo_id = 'MV1' and transaction_number='" & Star_No & "'"
'        End If
'    End If
'    RsUpdPro.Close: Set RsUpdPro = Nothing
End Sub


Private Sub CetakPromo(No_trans As String, Seq As Integer)
On Error GoTo ErrH
Dim RsPromo As New ADODB.Recordset, RsBayar As New ADODB.Recordset
Dim JmlKupon As Integer, NilaiOK As Long, Msg As String, ByrNonVoc As Long
Dim Msg1 As String
Dim Msg2 As String
Dim Msg3 As String
Dim Msg4 As String
Dim Gift_Voucher As Integer
Dim Nominal_Voucher As Long
Dim NilaiKK As Long
Dim NilaiKupon As Long, JmlKuponA As Integer, JmlKuponB As Integer, min_belanja As Long

    StrSQL = "SELECT promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn " & _
             ", SUM(Sales_Transaction_Details.Net_Price) As Belanja, islimit, qtylimit, qtyout, " & _
             " isnull(txt1,'') as txt1, isnull(txt2,'') as txt2, isnull(txt3,'') as txt3, isnull(txt4,'') as txt4, isnull(Gift_Voucher,0) as Gift_Voucher, isnull(Nominal,0) as Nominal FROM Promo_Hdr " & _
             "INNER JOIN Promo_Dtl ON Promo_Hdr.promo_id = Promo_Dtl.promo_id " & _
             "INNER JOIN Sales_Transaction_Details ON Promo_Dtl.PLU = Sales_Transaction_Details.PLU " & _
             "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
             "AND getdate() Between Start_Date And End_Date and aktif=1 And promo_hdr.seqmentation = " & Seq & "" & _
             "GROUP BY promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn, " & _
             "islimit, qtylimit, qtyout, txt1, txt2, txt3, txt4, Gift_Voucher,Nominal Having (promo_hdr.tipe > 30) and SUM(Sales_Transaction_Details.Net_Price)>0 " & _
             "order by promo_hdr.promo_id"
    
    If Linked Then
        RsPromo.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsPromo.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
    
    If RsPromo.EOF Then
        RsPromo.Close: Set RsPromo = Nothing
        Exit Sub
    End If

    RsBayar.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                 "From Paid where(Payment_Types not in ('8','5')) " & _
                 "GROUP BY Transaction_Number " & _
                 "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
    If Not RsBayar.EOF Then
        ByrNonVoc = RsBayar!bayar
    Else
        ByrNonVoc = 0
    End If
    RsBayar.Close: Set RsBayar = Nothing
    
    While Not RsPromo.EOF

        If Left(Star_Id, 6) = "100000" Or Star_Id = "" Then
            min_belanja = RsPromo!min_purchase
        Else
            min_belanja = RsPromo!min_member
        End If
        
        Msg = "": Msg1 = "": Msg2 = "": Msg3 = "": Msg4 = ""
        Gift_Voucher = 0
        Nominal_Voucher = 0
        Select Case RsPromo!Tipe
        Case 31, 32, 37, 40, 41, 43, 44 'GWP
            JmlKupon = 0
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then JmlKupon = 1
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                JmlKupon = roundDown(NilaiOK / min_belanja)
                If RsPromo!txt3 <> "" Then
                    If JmlKupon > RsPromo!txt3 Then JmlKupon = RsPromo!txt3
                End If
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!txt3 <> "" Then
                        If JmlKupon > RsPromo!txt3 Then JmlKupon = RsPromo!txt3
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
                
                
                 
                If RsPromo!Tipe = 31 Then 'GWP Normal
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                        Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!Gift_Voucher = 1 Then
                        Gift_Voucher = JmlKupon
                        Nominal_Voucher = RsPromo!Nominal
                    Else
                        Gift_Voucher = 0
                        Nominal_Voucher = 0
                    End If
                ElseIf RsPromo!Tipe = 32 Then 'Ultah STAR 300rb dpt Voucher 25rb kelipatan
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2 + vbNewLine
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(CLng(JmlKupon) * 25000, "#,##0")
                    Msg2 = "Rp " + Format(CLng(JmlKupon) * 25000, "#,##0")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg2 = Msg2 + " Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 37 Then 'GWP kupon disc% (tanpa qty pcs)
'                    Msg = "Tunjukan potongan struk ini dan dapatkan "
'                    Msg = Msg + vbNewLine + RsPromo!promo_name
'                    Msg = Msg + vbNewLine + "(+5% untuk Happy hour Jam 11-17, Senin - Kamis)"
'                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!Gift_Voucher = 1 Then
                        Gift_Voucher = JmlKupon
                        Nominal_Voucher = RsPromo!Nominal
                    Else
                        Gift_Voucher = 0
                        Nominal_Voucher = 0
                    End If
                ElseIf RsPromo!Tipe = 40 Then 'GWP khusus jumbo cash back
'                    If JmlKupon > 8 Then JmlKupon = 8
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg2 = Msg2 + " Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 44 Then 'GWP khusus jumbo cash back
'                    If JmlKupon > 8 Then JmlKupon = 8
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 60000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 60000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 60,000 = " & (JmlKupon * 60000) / 60000 & " Lembar"
                    Msg2 = Msg2 + "Voucher @ 60,000 = " & (JmlKupon * 60000) / 60000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 41 Then 'GWP khusus timezone
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name
'                    Msg = Msg + vbNewLine + "Syarat dan Ketentuan berlaku"
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku maksimal 7 hari"
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 43 Then 'GWP khusus jumbo cash back limit member
                    If Linked Then
                    Else
                        GoTo lanjut
                    End If
                    Dim JmlLimit As Integer
                    Dim RsBMember As New ADODB.Recordset
                    StrSQL = "Select b.ext1 As Potongan From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    RsBMember.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
                    If Not IsNumeric(RsBMember!Potongan) Then
                        JmlLimit = 0
                        GoTo lanjut
                    Else
                        If RsBMember!Potongan = 0 Then GoTo lanjut
                        JmlLimit = RsBMember!Potongan
                    End If
                    RsBMember.Close: Set RsBMember = Nothing
                    If JmlKupon > JmlLimit Then
                        JmlKupon = JmlLimit
                        ConnLocal.Execute "Update b set b.ext1 = 0 from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                        ConnServer.Execute "Update b set b.ext1 = 0 from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    Else
                        JmlLimit = JmlLimit - JmlKupon
                        ConnLocal.Execute "Update b set b.ext1 = " & JmlLimit & " from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                        ConnServer.Execute "Update b set b.ext1 = " & JmlLimit & " from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    End If
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg2 = Msg2 + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                 End If
                 
                 
                
                If Left(Star_No, 5) <> "CM000" Then Msg = Msg + vbNewLine + "MySTAR Card : " + Star_No
                If RsPromo!isprn = 1 Then
                    If StrukEmail = True Then
                        Call CetakStruk_PromoEmail(RsPromo!promo_id, No_trans, Msg1, Msg2, Msg3, Msg4, Gift_Voucher, Nominal_Voucher, Msg)
                    Else
                        Call CetakStruk_Promo(RsPromo!promo_id, No_trans, Gift_Voucher, Nominal_Voucher, Msg)
                    End If
                    
                End If
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
 
        Case 34 'GWP untuk SSC di hapus dulu
           
            
        Case 35 'GWP dan brand partisipasi (666) dapat tambahan 1 pcs
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                
                Dim RsLagi1 As New ADODB.Recordset
                
                RsLagi1.Open "SELECT Item_Master.Brand FROM Sales_Transaction_Details INNER JOIN " & _
                            "Promo_Dtl ON Sales_Transaction_Details.PLU = Promo_Dtl.PLU INNER JOIN " & _
                            "Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU " & _
                            "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
                            "AND (Promo_Dtl.promo_id = '666') GROUP BY Item_Master.Brand " & _
                            "HAVING SUM(Sales_Transaction_Details.Net_Price)>0", ConnLocal, adOpenForwardOnly, adLockReadOnly
                     
                If Not RsLagi1.EOF Then
                    JmlKupon = JmlKupon + 1
                End If
                RsLagi1.Close: Set RsLagi1 = Nothing
                
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")

                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    Else
                        
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"
                    MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut & " pcs", vbInformation, "Oops.."
                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
            End If

        Case 38 'GWP 500 ribu - 1 juta
            JmlKupon = 0
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If min_belanja = 1000000 Then
                    If NilaiOK >= 500000 And NilaiOK <= 999999 Then
                        JmlKupon = 1
                    End If
                End If
                
                If min_belanja = 750000 Then
                    If NilaiOK >= 750000 And NilaiOK <= 1499999 Then
                        JmlKupon = 1
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

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
                    Msg1 = "Anda mendapatkan :"
                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg2 = RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 39 'PWP 1 juta dan 2 juta
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= 1000000 And NilaiOK <= 1999999 Then
                JmlKupon = 1
            End If

            If NilaiOK >= 2000000 Then
                JmlKupon = 2
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt2
                    Msg1 = RsPromo!txt1 + " " + RsPromo!txt2
                    If JmlKupon = 2 Then
                        Msg = Msg + vbNewLine + RsPromo!txt3 + vbNewLine + RsPromo!txt4
                        Msg2 = RsPromo!txt3 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 42 'PWP Member bisa beli 2, non member bisa beli 1
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= min_belanja Then
                If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                    JmlKupon = 2
                Else
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    If JmlKupon = 1 Then
                        Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt2 + vbNewLine + RsPromo!txt4
                        Msg1 = RsPromo!txt1 + " " + RsPromo!txt2 + " " + RsPromo!txt4
                    Else
                        Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt3 + vbNewLine + RsPromo!txt4
                        Msg2 = RsPromo!txt1 + " " + RsPromo!txt3 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 50 'GWP for credit card
        
            JmlKupon = 0
            
            NilaiOK = Pakai_KK(RsPromo!promo_id, No_trans)
            
            If NilaiOK >= min_belanja Then
                If RsPromo!lipat = 0 Then
                    JmlKupon = 1
                ElseIf RsPromo!lipat = 1 Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                Else
                    If roundDown(NilaiOK / min_belanja) >= RsPromo!lipat Then
                    JmlKupon = RsPromo!lipat
                    Else
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
         
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                If RsPromo!tipe = 55 Then Msg = Msg + vbNewLine + "UNTUK 200 ORANG PENUKAR PERTAMA"
'                If RsPromo!tipe = 56 Then Msg = Msg + vbNewLine + "UNTUK 100 ORANG PENUKAR PERTAMA"
'                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 And StrukEmail = True Then Call CetakStruk_PromoEmail(RsPromo!promo_id, No_trans, Msg1, Msg2, Msg3, Msg4, Gift_Voucher, Nominal_Voucher, Msg)
                If RsPromo!isprn = 1 And StrukEmail = False Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, Gift_Voucher, Nominal_Voucher, Msg)

                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                'If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 51 'GWP lower purchase for bank
            JmlKupon = 0
            
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If

            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 52 'GWP Double untuk Kartu Bank
            JmlKupon = 0
                        
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            End If
            
            Dim JmlKuponTambah As Integer
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                If RsPromo!lipat = 1 Then
                    JmlKuponTambah = roundDown(Pakai_KK(RsPromo!promo_id, No_trans) / RsPromo!min_member)
                Else
                    JmlKuponTambah = 1
                End If
            End If
                
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
                
                    If RsPromo!txt1 <> "" Then Msg = RsPromo!txt1 + vbNewLine
                    If RsPromo!txt2 <> "" Then Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"

'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    If JmlKuponTambah > 0 Then
                        Msg = Msg + vbNewLine + "Tambahan = " & JmlKuponTambah & " pcs"
                        Msg = Msg + vbNewLine + "Total = " & JmlKupon + JmlKuponTambah & " pcs"
                        Msg1 = "Tambahan = " & JmlKuponTambah & " pcs"
                        Msg1 = Msg1 + " " + "Total = " & JmlKupon + JmlKuponTambah & " pcs"
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 53 'GWP hanya boleh 1 kali
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            End If
            
'            UNTUK PROMO MV1
'            Dim RsLagi2 As New ADODB.Recordset
'
'            StrSQL = "select * from promo_sales where Transaction_Number='" & Star_No & "' and promo_id='MV1'"
'
'            If Linked Then
'                RsLagi2.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
'            Else
'                RsLagi2.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
'            End If
'
'            If Not RsLagi2.EOF Then
'                JmlKupon = 0
'            End If
'            RsLagi2.Close: Set RsLagi1 = Nothing

            If JmlKupon > 0 Then
                If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If

        Case 58 'GWP 500 ribu - 1 juta bareng 31 tapi pakai_kk
            JmlKupon = 0
            
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If
            
            If ByrNonVoc < RsPromo!Belanja Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= min_belanja Then
                JmlKupon = 1
            End If
                
            If min_belanja = 1000000 Then

                If NilaiOK >= 500000 And NilaiOK <= 999999 Then
                    JmlKupon = 99
                End If
            End If
                
            If min_belanja = 1500000 Then
                If Pakai_KK(RsPromo!promo_id, No_trans) >= 500000 Then
                    JmlKupon = 99
                End If
                If NilaiOK >= 750000 And NilaiOK <= 1499999 Then
                    JmlKupon = 99
                End If
            End If
            
            If JmlKupon = 1 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        Msg = "Anda mendapatkan : " + vbNewLine + "Kesempatan membeli Teddy Bear Rp. 129,000 = " & JmlKupon & " pcs"
                        Msg1 = "Anda mendapatkan : " + vbNewLine + "Kesempatan membeli Teddy Bear Rp. 129,000 = " & JmlKupon & " pcs"
                        Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                        Msg2 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                        Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                        Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
            
                        If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                        If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                        If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"
                    MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut & " pcs", vbInformation, "Oops.."
                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    Msg = "Anda mendapatkan :"
                    Msg1 = "Anda mendapatkan :"
                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg2 = RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
            If JmlKupon = 99 Then
                Msg = "Anda mendapatkan : " + vbNewLine + "Kesempatan beli Teddy Bear Rp.129,000 = 1 pcs"
                Msg2 = "Anda mendapatkan : " + vbNewLine + "Kesempatan beli Teddy Bear Rp.129,000 = 1 pcs"
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
        
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
                GoTo lanjut
            End If
        Case 67 'Undian BNI
        JmlKupon = 0
            Dim NilaiCard As Long
            NilaiCard = Pakai_KK(RsPromo!promo_id, No_trans)
                If NilaiCard >= min_belanja Then
                    NilaiOK = NilaiCard
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            
            If JmlKupon > 0 Then
            Dim lagi2 As Integer
            
                For lagi2 = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                    
                   
                    
            If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
            If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
            Next lagi2
                
            If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 61, 62, 65 'Undian
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            
            If JmlKupon > 0 Then
            Dim lagi As Integer
            
                For lagi = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                    If RsPromo!Tipe = 65 Then
                        Dim NilaiKartu As Long
                        NilaiKartu = Pakai_KK(RsPromo!promo_id, No_trans)
                        If NilaiKartu > 0 Then
                            Msg = Msg + vbNewLine + "Pembayaran dgn Kartu Mandiri Rp." _
                                    + Format(NilaiKartu, "#,##0")
                        End If
                    End If
                    
            If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
            If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
            Next lagi
                
            If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 63 'Undian untuk item tertentu (special 17)
            Dim RsBayarAll As New ADODB.Recordset
            Dim ByrAll As Long
            
            RsBayarAll.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                            "From Paid GROUP BY Transaction_Number " & _
                            "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
            If Not RsBayarAll.EOF Then
                ByrAll = RsBayarAll!bayar
            Else
                ByrAll = 0
            End If
            RsBayarAll.Close: Set RsBayarAll = Nothing
    
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If ByrNonVoc >= min_belanja Then
                    NilaiOK = ByrNonVoc
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 Then
                If ByrAll >= min_belanja Then
                    NilaiOK = ByrAll
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                Msg = RsPromo!promo_name
                Msg1 = RsPromo!promo_name
                Msg = Msg + vbNewLine + "Nama :"
                Msg1 = Msg1 + vbNewLine + "Nama :"
                Msg = Msg + vbNewLine + vbNewLine + "No HP :"
                Msg2 = Msg2 + vbNewLine + vbNewLine + "No HP :"
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If


        Case 64, 66 'Undian dan brand partisipasi (666) dapat tambahan kupon
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            'Dobel Kupon
            If RsPromo!Tipe = 66 And Left(Star_No, 5) <> "CM000" Then JmlKupon = JmlKupon * 2
            
            If JmlKupon > 0 Then
            
            Dim RsLagi As New ADODB.Recordset
            
            RsLagi.Open "SELECT Item_Master.Brand FROM Sales_Transaction_Details INNER JOIN " & _
                        "Promo_Dtl ON Sales_Transaction_Details.PLU = Promo_Dtl.PLU INNER JOIN " & _
                        "Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU " & _
                        "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
                        "AND (Promo_Dtl.promo_id = '666') GROUP BY Item_Master.Brand " & _
                        "HAVING SUM(Sales_Transaction_Details.Net_Price)>0", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
            If Not RsLagi.EOF Then
                'JmlKupon = JmlKupon + RsLagi.RecordCount
                'perubahan 20 mei 2016
                JmlKupon = JmlKupon * 2
            End If
            RsLagi.Close: Set RsLagi = Nothing
            
            'Tambahan Kupon
            'If RsPromo!tipe = 66 And Left(Star_No, 5) <> "CM000" Then JmlKupon = JmlKupon + 1

            Dim lagi1 As Byte
            
                For lagi1 = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"

                    
                    If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                    If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                Next lagi1
                
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
         Case 71, 72, 73 'Voucher Progressif 300 rb = 30 rb dan 600 rb = 100 rb
            JmlKupon = 0: JmlKuponA = 0: JmlKuponB = 0
            Dim NilaiSatu As Double, NilaiDua As Double, NilaiPurchase As Double
            
            Select Case RsPromo!Tipe
                Case 71
                    NilaiSatu = 30000
                    NilaiDua = 100000
                    NilaiPurchase = 300000
                Case 72
                    NilaiSatu = 50000
                    NilaiDua = 150000
                    NilaiPurchase = 500000
                Case 73
                    NilaiSatu = 50000
                    NilaiDua = 150000
                    NilaiPurchase = 350000
            End Select
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= NilaiPurchase Then
                    If (NilaiOK \ NilaiPurchase) Mod 2 = 0 Then
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 0
                    Else
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua + NilaiSatu
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 1
                    End If
    
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    
                    If (NilaiOK \ NilaiPurchase) Mod 2 = 0 Then
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 0
                    Else
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua + NilaiSatu
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 1
                    End If
                End If
            End If
            
            If JmlKuponA > 0 Or JmlKuponB > 0 Then
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                    Msg1 = Msg1 + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                End If
                
                If RsPromo!Tipe = 73 Then
                    Msg = Msg + vbNewLine + "Voucher @ 50,000 = " & NilaiKupon / 50000 & " pcs"
                    Msg2 = "Voucher @ 50,000 = " & NilaiKupon / 50000 & " pcs"
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
    Case 74 'Voucher Progressif nilaiTrans=txt3 NilaiVouc=txt4 Contoh txt3 = 500,700,1000 dan txt4 = 30,50,80
        JmlKupon = 0: JmlKuponA = 0: JmlKuponB = 0
                Dim NilaiSatux As Double, NilaiDuax As Double, NilaiTigax As Double
                Dim NilaiTransSatu As Double, NilaiTransDua As Double, NilaiTransTiga As Double
            
                Select Case RsPromo!Tipe
                    Case 74
                    Dim NilaiVouc() As String
                      Dim I, j As Integer
                    Dim NilaiTrans() As String
                    NilaiTrans() = Split(RsPromo!txt3, ",")
                    For j = 0 To UBound(NilaiTrans)
                        Select Case j
                            Case 0
                            NilaiTransSatu = NilaiTrans(j) * 1000
                            Case 1
                            NilaiTransDua = NilaiTrans(j) * 1000
                            Case 2
                            NilaiTransTiga = NilaiTrans(j) * 1000
                        End Select
                    Next j
                    NilaiVouc() = Split(RsPromo!txt4, ",")
                    For I = 0 To UBound(NilaiVouc)
                        Select Case I
                            Case 0
                            NilaiSatux = NilaiVouc(I) * 1000
                            Case 1
                            NilaiDuax = NilaiVouc(I) * 1000
                            Case 2
                            NilaiTigax = NilaiVouc(I) * 1000
                        End Select
                    Next I
                End Select

                If RsPromo!voucher = 0 Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If

                If NilaiOK >= NilaiTransTiga Then
                    NilaiKupon = NilaiTigax
                ElseIf NilaiOK >= NilaiTransDua Then
                    NilaiKupon = NilaiDuax
                ElseIf NilaiOK >= NilaiTransSatu Then
                    NilaiKupon = NilaiSatux
                Else
                    NilaiKupon = 0
                End If

                If NilaiKupon > 0 Then
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                    Msg2 = RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                End If
                
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 75 'Progressif nilaiTrans=txt3 Info_Promo=txt4 Contoh txt3 = 500,700,1000 dan txt4 = 220,221,222 promo_id
     
                Select Case RsPromo!Tipe
                    Case 75
                    
                    Dim NilaiTrans75() As String
                    Dim NilaiInfo() As String
                    Dim NilaiTransx(10) As String
                    Dim NilaiInfox(10) As String
                    Dim k, l As Integer
                    
                    NilaiTrans75() = Split(RsPromo!txt3, ",")
                    For j = 0 To UBound(NilaiTrans75)
                        NilaiTransx(j) = NilaiTrans75(j) * 1000
                    Next j
                    NilaiInfo() = Split(RsPromo!txt4, ",")
                    For I = 0 To UBound(NilaiInfo)
                        NilaiInfox(I) = NilaiInfo(I)
                    Next I
                End Select
                If RsPromo!voucher = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                Else
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                End If
                End If
        

                PromoInfo = ""
                For j = UBound(NilaiTrans75) To 0 Step -1
                If NilaiOK >= NilaiTransx(j) Then
                    PromoInfo = NilaiInfox(j)
                    GoTo Loncat75
                End If
                Next j
Loncat75:

                If PromoInfo <> "" Then
                CekInfo75 (PromoInfo)
            
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 + PromoDesc
                    Msg2 = RsPromo!txt2 + PromoDesc
                End If
                
                
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                JmlKupon = 1
                If RsPromo!isprn = 1 Then
                    If StrukEmail = True Then
                        Call CetakStruk_PromoEmail(RsPromo!promo_id, No_trans, Msg1, Msg2, Msg3, Msg4, Gift_Voucher, Nominal_Voucher, Msg)
                    Else
                        Call CetakStruk_Promo(RsPromo!promo_id, No_trans, Gift_Voucher, Nominal_Voucher, Msg)
                    End If
                End If
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If

        End Select
lanjut:
    RsPromo.MoveNext
    Wend
    RsPromo.Close: Set RsPromo = Nothing
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Cetak Promo " & Err.Description & " " & Err.Number)
End Sub




Private Sub CetakPromo_Seq(No_trans As String, Seq As Integer, Promoid As String)
On Error GoTo ErrH
Dim RsPromo As New ADODB.Recordset, RsBayar As New ADODB.Recordset
Dim JmlKupon As Integer, NilaiOK As Long, Msg As String, ByrNonVoc As Long
Dim Msg1 As String
Dim Msg2 As String
Dim Msg3 As String
Dim Msg4 As String
Dim NilaiKK As Long
Dim NilaiKupon As Long, JmlKuponA As Integer, JmlKuponB As Integer, min_belanja As Long

    StrSQL = "SELECT promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn " & _
             ", SUM(Sales_Transaction_Details.Net_Price) As Belanja, islimit, qtylimit, qtyout, " & _
             " isnull(txt1,'') as txt1, isnull(txt2,'') as txt2, isnull(txt3,'') as txt3, isnull(txt4,'') as txt4 FROM Promo_Hdr " & _
             "INNER JOIN Promo_Dtl ON Promo_Hdr.promo_id = Promo_Dtl.promo_id " & _
             "INNER JOIN Sales_Transaction_Details ON Promo_Dtl.PLU = Sales_Transaction_Details.PLU " & _
             "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
             "AND getdate() Between Start_Date And End_Date and aktif=1 And promo_hdr.seqmentation <> 0 And promo_hdr.promo_id = '" & Promoid & "'" & _
             "GROUP BY promo_hdr.promo_id, promo_name, min_purchase, min_member, disc, tipe, voucher, lipat, ismsg, isprn, " & _
             "islimit, qtylimit, qtyout, txt1, txt2, txt3, txt4 Having (promo_hdr.tipe > 30) and SUM(Sales_Transaction_Details.Net_Price)>0 " & _
             "order by promo_hdr.promo_id"
    
    If Linked Then
        RsPromo.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    Else
        RsPromo.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    End If
    
    If RsPromo.EOF Then
        RsPromo.Close: Set RsPromo = Nothing
        UpdateStatusSeq = False
        Exit Sub
    End If

    RsBayar.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                 "From Paid where(Payment_Types not in ('8','5')) " & _
                 "GROUP BY Transaction_Number " & _
                 "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
    If Not RsBayar.EOF Then
        ByrNonVoc = RsBayar!bayar
    Else
        ByrNonVoc = 0
    End If
    RsBayar.Close: Set RsBayar = Nothing
    
    While Not RsPromo.EOF

        If Left(Star_Id, 6) = "100000" Or Star_Id = "" Then
            min_belanja = RsPromo!min_purchase
        Else
            min_belanja = RsPromo!min_member
        End If
        
        Msg = "": Msg1 = "": Msg2 = "": Msg3 = "": Msg4 = ""
        Select Case RsPromo!Tipe
        Case 31, 32, 37, 40, 41, 43, 44 'GWP
            JmlKupon = 0
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                JmlKupon = roundDown(NilaiOK / min_belanja)
                If RsPromo!txt3 <> "" Then
                    If JmlKupon > RsPromo!txt3 Then JmlKupon = RsPromo!txt3
                End If
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!txt3 <> "" Then
                        If JmlKupon > RsPromo!txt3 Then JmlKupon = RsPromo!txt3
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
                
                
                 
                If RsPromo!Tipe = 31 Then 'GWP Normal
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                        Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                ElseIf RsPromo!Tipe = 32 Then 'Ultah STAR 300rb dpt Voucher 25rb kelipatan
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2 + vbNewLine
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(CLng(JmlKupon) * 25000, "#,##0")
                    Msg2 = "Rp " + Format(CLng(JmlKupon) * 25000, "#,##0")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg2 = Msg2 + " Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 37 Then 'GWP kupon disc% (tanpa qty pcs)
'                    Msg = "Tunjukan potongan struk ini dan dapatkan "
'                    Msg = Msg + vbNewLine + RsPromo!promo_name
'                    Msg = Msg + vbNewLine + "(+5% untuk Happy hour Jam 11-17, Senin - Kamis)"
'                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                ElseIf RsPromo!Tipe = 40 Then 'GWP khusus jumbo cash back
'                    If JmlKupon > 8 Then JmlKupon = 8
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg2 = Msg2 + " Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 44 Then 'GWP khusus jumbo cash back
'                    If JmlKupon > 8 Then JmlKupon = 8
'                    Msg = "Anda mendapatkan STAR voucher senilai : "
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 60000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 60000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 60,000 = " & (JmlKupon * 60000) / 60000 & " Lembar"
                    Msg2 = Msg2 + "Voucher @ 60,000 = " & (JmlKupon * 60000) / 60000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 41 Then 'GWP khusus timezone
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name
'                    Msg = Msg + vbNewLine + "Syarat dan Ketentuan berlaku"
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku maksimal 7 hari"
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                ElseIf RsPromo!Tipe = 43 Then 'GWP khusus jumbo cash back limit member
                    If Linked Then
                    Else
                        GoTo lanjut
                    End If
                    Dim JmlLimit As Integer
                    Dim RsBMember As New ADODB.Recordset
                    StrSQL = "Select b.ext1 As Potongan From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    RsBMember.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
                    If Not IsNumeric(RsBMember!Potongan) Then
                        JmlLimit = 0
                        GoTo lanjut
                    Else
                        If RsBMember!Potongan = 0 Then GoTo lanjut
                        JmlLimit = RsBMember!Potongan
                    End If
                    RsBMember.Close: Set RsBMember = Nothing
                    If JmlKupon > JmlLimit Then
                        JmlKupon = JmlLimit
                        ConnLocal.Execute "Update b set b.ext1 = 0 from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                        ConnServer.Execute "Update b set b.ext1 = 0 from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    Else
                        JmlLimit = JmlLimit - JmlKupon
                        ConnLocal.Execute "Update b set b.ext1 = " & JmlLimit & " from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                        ConnServer.Execute "Update b set b.ext1 = " & JmlLimit & " from Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & Star_No & "'"
                    End If
                    If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    Msg = Msg + "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg2 = "Rp " + Format(JmlKupon * 200000, "#,##0")
                    Msg = Msg + vbNewLine + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg2 = Msg2 + "Voucher @ 100,000 = " & (JmlKupon * 200000) / 100000 & " Lembar"
                    Msg = Msg + vbNewLine + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg = Msg + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                 End If
                 
                 
                
                If Left(Star_No, 5) <> "CM000" Then Msg = Msg + vbNewLine + "MySTAR Card : " + Star_No
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
 
        Case 34 'GWP untuk SSC
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 And VIsSSC = True Then
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
            
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 35 'GWP dan brand partisipasi (666) dapat tambahan 1 pcs
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                
                Dim RsLagi1 As New ADODB.Recordset
                
                RsLagi1.Open "SELECT Item_Master.Brand FROM Sales_Transaction_Details INNER JOIN " & _
                            "Promo_Dtl ON Sales_Transaction_Details.PLU = Promo_Dtl.PLU INNER JOIN " & _
                            "Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU " & _
                            "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
                            "AND (Promo_Dtl.promo_id = '666') GROUP BY Item_Master.Brand " & _
                            "HAVING SUM(Sales_Transaction_Details.Net_Price)>0", ConnLocal, adOpenForwardOnly, adLockReadOnly
                     
                If Not RsLagi1.EOF Then
                    JmlKupon = JmlKupon + 1
                End If
                RsLagi1.Close: Set RsLagi1 = Nothing
                
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")

                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    Else
                        
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"
                    MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut & " pcs", vbInformation, "Oops.."
                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
            End If

        Case 38 'GWP 500 ribu - 1 juta
            JmlKupon = 0
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If min_belanja = 1000000 Then
                    If NilaiOK >= 500000 And NilaiOK <= 999999 Then
                        JmlKupon = 1
                    End If
                End If
                
                If min_belanja = 750000 Then
                    If NilaiOK >= 750000 And NilaiOK <= 1499999 Then
                        JmlKupon = 1
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

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
                    Msg1 = "Anda mendapatkan :"
                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg2 = RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 39 'PWP 1 juta dan 2 juta
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= 1000000 And NilaiOK <= 1999999 Then
                JmlKupon = 1
            End If

            If NilaiOK >= 2000000 Then
                JmlKupon = 2
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt2
                    Msg1 = RsPromo!txt1 + " " + RsPromo!txt2
                    If JmlKupon = 2 Then
                        Msg = Msg + vbNewLine + RsPromo!txt3 + vbNewLine + RsPromo!txt4
                        Msg2 = RsPromo!txt3 + vbNewLine + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 42 'PWP Member bisa beli 2, non member bisa beli 1
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= min_belanja Then
                If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                    JmlKupon = 2
                Else
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    If JmlKupon = 1 Then
                        Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt2 + vbNewLine + RsPromo!txt4
                        Msg1 = RsPromo!txt1 + " " + RsPromo!txt2 + " " + RsPromo!txt4
                    Else
                        Msg = RsPromo!txt1 + vbNewLine + RsPromo!txt3 + vbNewLine + RsPromo!txt4
                        Msg2 = RsPromo!txt1 + " " + RsPromo!txt3 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
        Case 50 'GWP for credit card
        
            JmlKupon = 0
            
            NilaiOK = Pakai_KK(RsPromo!promo_id, No_trans)
            
            If NilaiOK >= min_belanja Then
                If RsPromo!lipat = 0 Then
                    JmlKupon = 1
                ElseIf RsPromo!lipat = 1 Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                Else
                    If roundDown(NilaiOK / min_belanja) >= RsPromo!lipat Then
                    JmlKupon = RsPromo!lipat
                    Else
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    End If
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
         
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                If RsPromo!tipe = 55 Then Msg = Msg + vbNewLine + "UNTUK 200 ORANG PENUKAR PERTAMA"
'                If RsPromo!tipe = 56 Then Msg = Msg + vbNewLine + "UNTUK 100 ORANG PENUKAR PERTAMA"
'                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                    
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 51 'GWP lower purchase for bank
            JmlKupon = 0
            
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If

            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            End If
            
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
'                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
'                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                    Msg1 = Msg1 + RsPromo!txt2 & " = " & JmlKupon & " pcs"
                End If
                If RsPromo!txt3 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt3
                    Msg2 = RsPromo!txt3
                End If
                If RsPromo!txt4 <> "" Then
                    Msg = Msg + vbNewLine + RsPromo!txt4
                    Msg2 = Msg2 + " " + RsPromo!txt4
                End If
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 52 'GWP Double untuk Kartu Bank
            JmlKupon = 0
                        
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 0 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= min_belanja Then
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                End If
            End If
            
            Dim JmlKuponTambah As Integer
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                If RsPromo!lipat = 1 Then
                    JmlKuponTambah = roundDown(Pakai_KK(RsPromo!promo_id, No_trans) / RsPromo!min_member)
                Else
                    JmlKuponTambah = 1
                End If
            End If
                
            If JmlKupon > 0 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        MsgBox "Hadiah " & RsPromo!promo_name & " sudah habis", vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"

                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                End If
                
                    If RsPromo!txt1 <> "" Then Msg = RsPromo!txt1 + vbNewLine
                    If RsPromo!txt2 <> "" Then Msg = Msg + RsPromo!txt2 & " = " & JmlKupon & " pcs"

'                    Msg = "Anda mendapatkan :"
'                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    If JmlKuponTambah > 0 Then
                        Msg = Msg + vbNewLine + "Tambahan = " & JmlKuponTambah & " pcs"
                        Msg = Msg + vbNewLine + "Total = " & JmlKupon + JmlKuponTambah & " pcs"
                        Msg1 = "Tambahan = " & JmlKuponTambah & " pcs"
                        Msg1 = Msg1 + " " + "Total = " & JmlKupon + JmlKuponTambah & " pcs"
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + " " + RsPromo!txt4
                    End If
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 53 'GWP hanya boleh 1 kali
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = 1
                End If
            End If
            
'            UNTUK PROMO MV1
'            Dim RsLagi2 As New ADODB.Recordset
'
'            StrSQL = "select * from promo_sales where Transaction_Number='" & Star_No & "' and promo_id='MV1'"
'
'            If Linked Then
'                RsLagi2.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
'            Else
'                RsLagi2.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
'            End If
'
'            If Not RsLagi2.EOF Then
'                JmlKupon = 0
'            End If
'            RsLagi2.Close: Set RsLagi1 = Nothing

            If JmlKupon > 0 Then
                If RsPromo!txt1 <> "" Then
                        Msg = RsPromo!txt1 + vbNewLine
                        Msg1 = RsPromo!txt1
                    End If
                    If RsPromo!txt2 <> "" Then
                        Msg = Msg + RsPromo!txt2
                        Msg1 = Msg1 + RsPromo!txt2
                    End If
                    If RsPromo!txt3 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt3
                        Msg2 = RsPromo!txt3
                    End If
                    If RsPromo!txt4 <> "" Then
                        Msg = Msg + vbNewLine + RsPromo!txt4
                        Msg2 = Msg2 + vbNewLine + RsPromo!txt4
                    End If
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If

        Case 58 'GWP 500 ribu - 1 juta bareng 31 tapi pakai_kk
            JmlKupon = 0
            
            If Pakai_KK(RsPromo!promo_id, No_trans) >= RsPromo!min_member Then
                min_belanja = RsPromo!min_member
            End If
            
            If ByrNonVoc < RsPromo!Belanja Then
                NilaiOK = ByrNonVoc
            Else
                NilaiOK = RsPromo!Belanja
            End If
                
            If NilaiOK >= min_belanja Then
                JmlKupon = 1
            End If
                
            If min_belanja = 1000000 Then

                If NilaiOK >= 500000 And NilaiOK <= 999999 Then
                    JmlKupon = 99
                End If
            End If
                
            If min_belanja = 1500000 Then
                If Pakai_KK(RsPromo!promo_id, No_trans) >= 500000 Then
                    JmlKupon = 99
                End If
                If NilaiOK >= 750000 And NilaiOK <= 1499999 Then
                    JmlKupon = 99
                End If
            End If
            
            If JmlKupon = 1 Then
                If RsPromo!islimit = 1 Then
                    If RsPromo!QtyLimit < RsPromo!QtyOut + JmlKupon Then
                        Msg = "Anda mendapatkan : " + vbNewLine + "Kesempatan membeli Teddy Bear Rp. 129,000 = " & JmlKupon & " pcs"
                        Msg1 = "Anda mendapatkan : " + vbNewLine + "Kesempatan membeli Teddy Bear Rp. 129,000 = " & JmlKupon & " pcs"
                        Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                        Msg2 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                        Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                        Msg3 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
            
                        If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                        If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                        If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                        GoTo lanjut
                    End If
                    
                    StrSQL = "insert into promo_sales(promo_id,transaction_number,nilai,qty_promo,status) values ('" & _
                             RsPromo!promo_id & "', '" & No_trans & "', " & NilaiOK & ", " & JmlKupon & ", '00')"

                    ConnLocal.Execute StrSQL
                    If Linked Then
                        ConnServer.Execute StrSQL
                        ConnLocal.Execute "Update promo_sales set status='99' where promo_id = '" & RsPromo!promo_id & "' and transaction_number='" & No_trans & "'"
                    End If
                    
                    StrSQL = "update promo_hdr set qtyout=qtyout+ " & JmlKupon & " where promo_id='" & RsPromo!promo_id & "'"
                    MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut & " pcs", vbInformation, "Oops.."
                    ConnLocal.Execute StrSQL
                    If Linked Then ConnServer.Execute StrSQL
                    
                End If
                
                    Msg = "Anda mendapatkan :"
                    Msg1 = "Anda mendapatkan :"
                    Msg = Msg + vbNewLine + RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg2 = RsPromo!promo_name & " = " & JmlKupon & " pcs"
                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
            If JmlKupon = 99 Then
                Msg = "Anda mendapatkan : " + vbNewLine + "Kesempatan beli Teddy Bear Rp.129,000 = 1 pcs"
                Msg2 = "Anda mendapatkan : " + vbNewLine + "Kesempatan beli Teddy Bear Rp.129,000 = 1 pcs"
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
        
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
                GoTo lanjut
            End If
        Case 67 'Undian BNI
        JmlKupon = 0
            Dim NilaiCard As Long
            NilaiCard = Pakai_KK(RsPromo!promo_id, No_trans)
                If NilaiCard >= min_belanja Then
                    NilaiOK = NilaiCard
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            
            If JmlKupon > 0 Then
            Dim lagi2 As Integer
            
                For lagi2 = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                    
                   
                    
            If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
            
            Next lagi2
                
            If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 61, 62, 65 'Undian
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            
            If JmlKupon > 0 Then
            Dim lagi As Integer
            
                For lagi = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                    
                    If RsPromo!Tipe = 65 Then
                        Dim NilaiKartu As Long
                        NilaiKartu = Pakai_KK(RsPromo!promo_id, No_trans)
                        If NilaiKartu > 0 Then
                            Msg = Msg + vbNewLine + "Pembayaran dgn Kartu Mandiri Rp." _
                                    + Format(NilaiKartu, "#,##0")
                        End If
                    End If
                    
            If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
            If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
            Next lagi
                
            If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
            If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 63 'Undian untuk item tertentu (special 17)
            Dim RsBayarAll As New ADODB.Recordset
            Dim ByrAll As Long
            
            RsBayarAll.Open "SELECT Transaction_Number, SUM(Paid_Amount) AS Bayar " & _
                            "From Paid GROUP BY Transaction_Number " & _
                            "HAVING (Transaction_Number = '" & No_trans & "')", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
            If Not RsBayarAll.EOF Then
                ByrAll = RsBayarAll!bayar
            Else
                ByrAll = 0
            End If
            RsBayarAll.Close: Set RsBayarAll = Nothing
    
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If ByrNonVoc >= min_belanja Then
                    NilaiOK = ByrNonVoc
                    JmlKupon = 1
                End If
            ElseIf RsPromo!voucher = 1 Then
                If ByrAll >= min_belanja Then
                    NilaiOK = ByrAll
                    JmlKupon = 1
                End If
            End If
            
            If JmlKupon > 0 Then
                Msg = RsPromo!promo_name
                Msg1 = RsPromo!promo_name
                Msg = Msg + vbNewLine + "Nama :"
                Msg1 = Msg1 + vbNewLine + "Nama :"
                Msg = Msg + vbNewLine + vbNewLine + "No HP :"
                Msg2 = Msg2 + vbNewLine + vbNewLine + "No HP :"
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If


        Case 64, 66 'Undian dan brand partisipasi (666) dapat tambahan kupon
            JmlKupon = 0
            
            If RsPromo!voucher = 0 Then
                If RsPromo!Belanja < ByrNonVoc Then
                    If RsPromo!Belanja >= min_belanja Then
                        NilaiOK = RsPromo!Belanja
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                Else
                    If ByrNonVoc >= min_belanja Then
                        NilaiOK = ByrNonVoc
                        JmlKupon = roundDown(NilaiOK / min_belanja)
                        If RsPromo!lipat = 0 Then JmlKupon = 1
                    End If
                End If
            ElseIf RsPromo!voucher = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    JmlKupon = roundDown(NilaiOK / min_belanja)
                    If RsPromo!lipat = 0 Then JmlKupon = 1
                End If
            End If
            
            If Left(Star_No, 5) = "CM999" Then JmlKupon = 0
            'Dobel Kupon
            If RsPromo!Tipe = 66 And Left(Star_No, 5) <> "CM000" Then JmlKupon = JmlKupon * 2
            
            If JmlKupon > 0 Then
            
            Dim RsLagi As New ADODB.Recordset
            
            RsLagi.Open "SELECT Item_Master.Brand FROM Sales_Transaction_Details INNER JOIN " & _
                        "Promo_Dtl ON Sales_Transaction_Details.PLU = Promo_Dtl.PLU INNER JOIN " & _
                        "Item_Master ON Sales_Transaction_Details.PLU = Item_Master.PLU " & _
                        "WHERE (Sales_Transaction_Details.Transaction_Number = '" & No_trans & "') " & _
                        "AND (Promo_Dtl.promo_id = '666') GROUP BY Item_Master.Brand " & _
                        "HAVING SUM(Sales_Transaction_Details.Net_Price)>0", ConnLocal, adOpenForwardOnly, adLockReadOnly
                 
            If Not RsLagi.EOF Then
                'JmlKupon = JmlKupon + RsLagi.RecordCount
                'perubahan 20 mei 2016
                JmlKupon = JmlKupon * 2
            End If
            RsLagi.Close: Set RsLagi = Nothing
            
            'Tambahan Kupon
            'If RsPromo!tipe = 66 And Left(Star_No, 5) <> "CM000" Then JmlKupon = JmlKupon + 1

            Dim lagi1 As Byte
            
                For lagi1 = 1 To JmlKupon
                    Msg = RsPromo!promo_name & " #" & lagi2
                    Msg1 = RsPromo!promo_name & " #" & lagi2
                    If Left(Star_No, 5) <> "CM000" And Star_Id <> "100000" Then
                        Msg = Msg + vbNewLine + "Nama : " & Star_Nm
                        Msg1 = Msg1 + vbNewLine + "Nama : " & Star_Nm
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : " & Star_Phone
                        Msg2 = "No HP : " & Star_Phone
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : " & Star_Id
                    Else
                        Msg = Msg + vbNewLine + "Nama : "
                        Msg1 = Msg1 + vbNewLine + "Nama : "
                        Msg = Msg + vbNewLine + vbNewLine + "No HP : "
                        Msg2 = "No HP : "
                        Msg = Msg + vbNewLine + vbNewLine + "No KTP : "
                        Msg2 = Msg2 + vbNewLine + vbNewLine + "No KTP : "
                    End If
                    
'                    Msg = Msg + vbNewLine + "Nama :"
'                    Msg = Msg + vbNewLine + vbNewLine + "No HP :"
'                    If RsPromo!tipe = 62 Then Msg = Msg + vbNewLine + vbNewLine + "No KTP :"

                    Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                    Msg = Msg + vbNewLine + "Masukan struk ini di Kotak Undian"
                    Msg3 = Msg3 + vbNewLine + "Masukan struk ini di Kotak Undian"

                    
                    If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                    
                Next lagi1
                
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
         Case 71, 72, 73 'Voucher Progressif 300 rb = 30 rb dan 600 rb = 100 rb
            JmlKupon = 0: JmlKuponA = 0: JmlKuponB = 0
            Dim NilaiSatu As Double, NilaiDua As Double, NilaiPurchase As Double
            
            Select Case RsPromo!Tipe
                Case 71
                    NilaiSatu = 30000
                    NilaiDua = 100000
                    NilaiPurchase = 300000
                Case 72
                    NilaiSatu = 50000
                    NilaiDua = 150000
                    NilaiPurchase = 500000
                Case 73
                    NilaiSatu = 50000
                    NilaiDua = 150000
                    NilaiPurchase = 350000
            End Select
            
            If RsPromo!voucher = 0 And RsPromo!lipat = 1 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                
                If NilaiOK >= NilaiPurchase Then
                    If (NilaiOK \ NilaiPurchase) Mod 2 = 0 Then
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 0
                    Else
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua + NilaiSatu
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 1
                    End If
    
                End If
            ElseIf RsPromo!voucher = 1 And RsPromo!lipat = 1 Then
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                    
                    If (NilaiOK \ NilaiPurchase) Mod 2 = 0 Then
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 0
                    Else
                        NilaiKupon = (NilaiOK \ (NilaiPurchase * 2)) * NilaiDua + NilaiSatu
                        JmlKuponA = (NilaiOK \ (NilaiPurchase * 2))
                        JmlKuponB = 1
                    End If
                End If
            End If
            
            If JmlKuponA > 0 Or JmlKuponB > 0 Then
'                Msg = "Anda mendapatkan :"
'                Msg = Msg + vbNewLine + RsPromo!promo_name & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                    Msg1 = Msg1 + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                End If
                
                Select Case RsPromo!Tipe
'                Case 15
'                    If JmlKuponA <> 0 Then Msg = Msg + vbNewLine + "Voucher @ 100,000 = " & JmlKuponA & " pcs"
'                    If JmlKuponB <> 0 Then Msg = Msg + vbNewLine + "Voucher @ 30,000 = " & JmlKuponB & " pcs"
'                Case 16
'                    If JmlKuponA <> 0 Then Msg = Msg + vbNewLine + "Voucher @ 150,000 = " & JmlKuponA & " pcs"
'                    If JmlKuponB <> 0 Then Msg = Msg + vbNewLine + "Voucher @ 50,000 = " & JmlKuponB & " pcs"
                Case 73
                    Msg = Msg + vbNewLine + "Voucher @ 50,000 = " & NilaiKupon / 50000 & " pcs"
                    Msg2 = "Voucher @ 50,000 = " & NilaiKupon / 50000 & " pcs"
                End Select
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
    Case 74 'Voucher Progressif nilaiTrans=txt3 NilaiVouc=txt4 Contoh txt3 = 500,700,1000 dan txt4 = 30,50,80
        JmlKupon = 0: JmlKuponA = 0: JmlKuponB = 0
                Dim NilaiSatux As Double, NilaiDuax As Double, NilaiTigax As Double
                Dim NilaiTransSatu As Double, NilaiTransDua As Double, NilaiTransTiga As Double
            
                Select Case RsPromo!Tipe
                    Case 74
                    Dim NilaiVouc() As String
                      Dim I, j As Integer
                    Dim NilaiTrans() As String
                    NilaiTrans() = Split(RsPromo!txt3, ",")
                    For j = 0 To UBound(NilaiTrans)
                        Select Case j
                            Case 0
                            NilaiTransSatu = NilaiTrans(j) * 1000
                            Case 1
                            NilaiTransDua = NilaiTrans(j) * 1000
                            Case 2
                            NilaiTransTiga = NilaiTrans(j) * 1000
                        End Select
                    Next j
                    NilaiVouc() = Split(RsPromo!txt4, ",")
                    For I = 0 To UBound(NilaiVouc)
                        Select Case I
                            Case 0
                            NilaiSatux = NilaiVouc(I) * 1000
                            Case 1
                            NilaiDuax = NilaiVouc(I) * 1000
                            Case 2
                            NilaiTigax = NilaiVouc(I) * 1000
                        End Select
                    Next I
                End Select

                If RsPromo!voucher = 0 Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If

                If NilaiOK >= NilaiTransTiga Then
                    NilaiKupon = NilaiTigax
                ElseIf NilaiOK >= NilaiTransDua Then
                    NilaiKupon = NilaiDuax
                ElseIf NilaiOK >= NilaiTransSatu Then
                    NilaiKupon = NilaiSatux
                Else
                    NilaiKupon = 0
                End If

                If NilaiKupon > 0 Then
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                    Msg2 = RsPromo!txt2 & " Senilai = Rp. " & Format(NilaiKupon, "#,##0")
                End If
                
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If
            
        Case 75 'Progressif nilaiTrans=txt3 Info_Promo=txt4 Contoh txt3 = 500,700,1000 dan txt4 = 220,221,222 promo_id
        
                'Dim NilaiSatux As Double, NilaiDuax As Double, NilaiTigax As Double
                'Dim NilaiTransSatu As Double, NilaiTransDua As Double, NilaiTransTiga As Double
            
                Select Case RsPromo!Tipe
                    Case 75
                    
                    Dim NilaiTrans75() As String
                    Dim NilaiInfo() As String
                    Dim NilaiTransx(10) As String
                    Dim NilaiInfox(10) As String
                    Dim k, l As Integer
                    
                    NilaiTrans75() = Split(RsPromo!txt3, ",")
                    For j = 0 To UBound(NilaiTrans75)
                        NilaiTransx(j) = NilaiTrans75(j) * 1000
                    Next j
                    NilaiInfo() = Split(RsPromo!txt4, ",")
                    For I = 0 To UBound(NilaiInfo)
                        NilaiInfox(I) = NilaiInfo(I)
                    Next I
                End Select
                If RsPromo!voucher = 0 Then
                If ByrNonVoc < RsPromo!Belanja Then
                    NilaiOK = ByrNonVoc
                Else
                    NilaiOK = RsPromo!Belanja
                End If
                Else
                If RsPromo!Belanja >= min_belanja Then
                    NilaiOK = RsPromo!Belanja
                End If
                End If
        

                PromoInfo = ""
                For j = UBound(NilaiTrans75) To 0 Step -1
                If NilaiOK >= NilaiTransx(j) Then
                    PromoInfo = NilaiInfox(j)
                    GoTo Loncat75
                End If
                Next j
Loncat75:

                If PromoInfo <> "" Then
                CekInfo75 (PromoInfo)
            
                If RsPromo!txt1 <> "" Then
                    Msg = RsPromo!txt1 + vbNewLine
                    Msg1 = RsPromo!txt1
                End If
                If RsPromo!txt2 <> "" Then
                    Msg = Msg + RsPromo!txt2 + PromoDesc
                    Msg2 = RsPromo!txt2 + PromoDesc
                End If
                
                
                Msg = Msg + vbNewLine + vbNewLine + "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg3 = "Nilai Transaksi : Rp." + Format(NilaiOK, "#,##0")
                Msg = Msg + vbNewLine + "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                Msg4 = "Struk ini hanya berlaku tgl " & Format(GetSrvDate(), "DD mmm 'YY")
                JmlKupon = 1
                If RsPromo!isprn = 1 Then Call CetakStruk_Promo(RsPromo!promo_id, No_trans, 0, 0, Msg)
                If RsPromo!isprn = 2 Then Call Kirim_Promo_Mobile(RsPromo!promo_id, 0, Star_No, No_trans, Msg1, Msg2, Msg3, Msg4, JmlKupon, Msg)
                If RsPromo!ismsg = 1 Then MsgBox Msg, vbInformation, "Oops.."
                If RsPromo!islimit = 1 Then MsgBox "Sisa Stok Hadiah " & RsPromo!promo_name & " " & RsPromo!QtyLimit - RsPromo!QtyOut - 1 & " pcs", vbInformation, "Oops.."
            End If

        End Select
lanjut:
    RsPromo.MoveNext
    Wend
    RsPromo.Close: Set RsPromo = Nothing
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Cek Lunas " & Err.Description & " " & Err.Number)
End Sub

Private Function CekInfo75(ProgId As String) As Long
                Dim CekInfox As New ADODB.Recordset
    
                CekInfox.Open "SELECT * from promo_dtl where promo_id = '" & ProgId & "'", ConnLocal, adOpenForwardOnly, adLockReadOnly
                    
                If Not CekInfox.EOF Then
                    PromoDesc = CekInfox!Description
                Else
                    PromoDesc = ""
                End If
                CekInfox.Close: Set CekInfox = Nothing
End Function

Private Function Pakai_KK(ProgId As String, No_trans As String) As Long
Dim RsKartuKredit As New ADODB.Recordset
    
    StrSQL = "select isnull(sum(paid_amount),0) as NilaiKK from paid where transaction_number='" & _
    No_trans & "' and left(credit_card_no,6) in (select CAST(nomor AS varchar(6)) from cc_master where cc_master='" & ProgId & "')"
    RsKartuKredit.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    Pakai_KK = RsKartuKredit!NilaiKK
    RsKartuKredit.Close: Set RsKartuKredit = Nothing
End Function

'Private Function Pakai_KK2(ProgId As String, No_trans As String) As Long
'Dim RsKartuKredit As New ADODB.Recordset
    
'StrSQL = "select isnull(sum(net_price),0) as NilaiKK from Sales_Transaction_Details  where transaction_number='" & No_trans & "'"
'RsKartuKredit.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
'Pakai_KK2 = RsKartuKredit!NilaiKK
'RsKartuKredit.Close: Set RsKartuKredit = Nothing
'End Function


Private Function Cek_BonusKaryawan(NoKartu As String, No_trans As String) As Long
Dim RsBKaryawan As New ADODB.Recordset
    
    StrSQL = "Select b.ext1 As Potongan From Card a inner join Customer_Master_Member b on a.Cust_Nr = b.Cust_Nr Where a.Card_Nr = '" & NoKartu & "'"
    RsBKaryawan.Open StrSQL, ConnServer, adOpenForwardOnly, adLockReadOnly
    If Not IsNumeric(RsBKaryawan!Potongan) Then
    Cek_BonusKaryawan = 0
    Else
    Cek_BonusKaryawan = RsBKaryawan!Potongan
    If RsBKaryawan!Potongan <= 10 Then
    Cek_BonusKaryawan = 0
    End If
    End If
    RsBKaryawan.Close: Set RsBKaryawan = Nothing
    If Cek_BonusKaryawan = 0 Then
    StrSQL = "select * from paid where transaction_number='" & No_trans & "' and Payment_Types = '31'"
    RsBKaryawan.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly
    If RsBKaryawan.RecordCount > 0 Then
    Cek_BonusKaryawan = 1
    Else
    Cek_BonusKaryawan = 0
    End If
    End If
End Function

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
            ", total_discount=" & RsHitung!hemat & ", status='00' , net_price=net_amount , transaction_time='" & Format(GetSrvDate, "HH:NN") & "' where transaction_number = '" & nomor & "'"
    
    ConnLocal.Execute StrSQL
    
    RsHitung.Close: Set RsHitung = Nothing
    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Paid_To_Sales " & Err.Description & " " & Err.Number)
End Sub

Private Sub Upload_to_Server(nomor As String)
On Error GoTo ErrH
'On Error Resume Next
Dim Dbs As String, Svr As String

    Svr = "[" & VSvr & "]"
    Dbs = Cfg_Get("Server", "DatabaseName", App.Path & "\config.ini")
    ''apabila tombol exis enable saat hold transaksi
    ''StrSQL = "delete from " & Svr & "." & Dbs & ".dbo.Sales_Transactions where Transaction_Number='" & nomor & "'"
    ''ConnLocal.Execute StrSQL
    ''StrSQL = "delete from " & Svr & "." & Dbs & ".dbo.Sales_Transaction_details where Transaction_Number='" & nomor & "'"
    ''ConnLocal.Execute StrSQL
    ''StrSQL = "delete from " & Svr & "." & Dbs & ".dbo.paid where Transaction_Number='" & nomor & "'"
    ''ConnLocal.Execute StrSQL
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

    StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.Sales_Transaction_details (Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
        "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, " & _
        "Flag_Status, Flag_Paket_Discount) " & _
        "(select Transaction_Number, Seq, PLU, Item_Description, Price, Qty, Amount, " & _
        "Discount_Percentage, Discount_Amount, ExtraDisc_Pct, ExtraDisc_Amt, Net_Price, Points_Received, Flag_Void, " & _
        "Flag_Status , Flag_Paket_Discount " & _
        "FROM Sales_Transaction_details where Transaction_Number='" & nomor & "')"

    ConnLocal.Execute StrSQL
    
    StrSQL = "Insert " & Svr & "." & Dbs & ".dbo.paid(Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, " & _
        "Credit_Card_No, Credit_Card_Name, Paid_Amount, Shift) " & _
        "(select Transaction_Number, Payment_Types, Seq, Currency_ID, Currency_Rate, Credit_Card_No, " & _
        "Credit_Card_Name , Paid_Amount, Shift " & _
        "FROM paid where Transaction_Number='" & nomor & "')"
    
    ConnLocal.Execute StrSQL
    
    ConnLocal.Execute "Update sales_transactions set upload_status='99' where transaction_number='" & nomor & "'"

    Exit Sub
    
ErrH:
    MsgBox UCase(Err.Description), vbCritical + vbOKOnly, "Oops.."
    Call SaveLog(Me.name & " " & "Upload_to_Server " & Err.Description & " " & Err.Number)
End Sub

Private Sub tampil_point()
    txtsaldo_point = Star_Pt
    
    'If (lbltotal * 0.2 / txtharga_point) < txtsaldo_point Then
    '    txtpoint = roundDown(lbltotal * 0.2)
    'Else
        txtpoint = txtsaldo_point * txtharga_point
    'End If
    
    txttukar_point = Format(roundDown(txtpoint / txtharga_point), "#,##0")
End Sub

Private Sub txtpoint_GotFocus()
    lokasi = "txtpoint"
End Sub

Private Sub txttukar_point_GotFocus()
    lokasi = "txttukar_point"
End Sub

Private Sub txttukar_point_Change()
On Error GoTo x
    txtpoint = Format(txttukar_point * txtharga_point, "#,###,##0")
    Exit Sub
x:
    txttukar_point = txtsaldo_point
    txtpoint = Format(txttukar_point * txtharga_point, "#,###,##0")
End Sub

Private Sub txtpoint_lostfocus()
    txttukar_point = Format(roundDown(txtpoint / txtharga_point), "#,###,##0")
    txtpoint = Format(txttukar_point * txtharga_point, "#,###,##0")
End Sub

Private Sub txtpoint_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            txttukar_point.SetFocus
        Case 27
            GridPay_Types.SetFocus
    End Select
End Sub

Private Sub txttukar_point_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13

        If Linked = False Then
            MsgBox "Pembayaran dengan point reward harus Online", vbOKOnly + vbInformation, "Oops.."
            Exit Sub
        End If
        
        If txtsaldo_point.Value < 20 Then
            MsgBox "Minimal Saldo point harus 20", vbOKOnly + vbInformation, "Oops.."
            Exit Sub
        End If
        
        If txttukar_point.Value > txtsaldo_point Then
            MsgBox "Saldo point tidak mencukupi", vbOKOnly + vbInformation, "Oops.."
            Exit Sub
        End If
        
        'If txtpoint.Value / lbltotal > 0.2 Then
        '    MsgBox "Pembayaran dengan point max 20% dari nilai transaksi", vbOKOnly + vbInformation, "Oops.."
        '   Exit Sub
        'End If
        
        If txtcard_no = "CM000-00000" Then
           MsgBox "Pembayaran dengan point hanya untuk member MSC", vbOKOnly + vbInformation, "Oops.."
           Exit Sub
        End If
        
        Dim RsPr As New ADODB.Recordset
            StrSQL = "SELECT * from paid WHERE payment_types='5' and transaction_number='" & vno_trans & "'"
            RsPr.Open StrSQL, ConnLocal, adOpenForwardOnly, adLockReadOnly

        If Not RsPr.EOF Then
            MsgBox "Pembayaran dengan point hanya bisa 1 kali/transaksi", vbOKOnly + vbInformation, "Oops.."
            RsPr.Close: Set RsPr = Nothing
            Exit Sub
        End If
        RsPr.Close: Set RsPr = Nothing
            
        If txttukar_point.Value > 0 Then
            If txtpoint.Value > vpay Then
                MsgBox "Pembayaran dengan point reward tidak boleh " & vbNewLine & _
                "melebihi sisa yang harus dibayar", vbOKOnly + vbInformation, "Oops.."
                txttukar_point.SetFocus
                DoEvents
                SendKeys "{home}+{end}"
            Else
                GridPay_Types.SetFocus
                VTukar_Point = Pay_Point(txttukar_point, txtcard_no, vno_trans, txtpoint)
                If VTukar_Point <> "GAGAL" Then
                    Call Simpan_Detail(txtpoint, txtcard_no, VTukar_Point)
                    Call tampil_point
                    txttukar_point = 0
                    Call Cek_Lunas
                Else
                    MsgBox "Pembayaran dengan point reward GAGAL", vbOKOnly + vbInformation, "Oops.."
                    Exit Sub
                End If
            End If
        Else
            MsgBox "Transaksi Refund tidak bisa menggunakan point reward", vbOKOnly + vbInformation, "Oops.."
        End If
    Case 27
        GridPay_Types.SetFocus
    End Select
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

Private Sub btnNum_Click(Index As Integer)
On Error Resume Next
Dim box As Control

Set box = Me.Controls(lokasi)
    Select Case lokasi
    Case "txtcash", "txtcredit", "txtvoucher", "txttukar_point"
        If Index < 10 Then box.Value = box.Text + CStr(btnNum(Index).Caption)
        Select Case Index
        Case 10
            box.SetFocus
            SendKeys "{end}+{backspace}"
        Case 11
            box.SetFocus
            SendKeys "{enter}"
        Case 12
            box.Value = 0
        End Select
    Case "txtno_kartu", "txtno_voc"
        If Index < 10 Then box.Text = box.Text + CStr(btnNum(Index).Caption)

        Select Case Index
        Case 10
            box.SetFocus
            SendKeys "{end}+{backspace}"
        Case 11
            box.SetFocus
            SendKeys "{enter}"
        Case 12
            box.Text = ""
        End Select
    End Select
    
    box.SetFocus
    SendKeys "{end}"
End Sub
