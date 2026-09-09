VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.MDIForm EmrFrmJMDI 
   BackColor       =   &H00808080&
   Caption         =   "jEmr"
   ClientHeight    =   3315
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   5235
   Icon            =   "EmrFrmJMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "EmrFrmJMDI.frx":1272
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   5235
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   5235
      Begin VB.ComboBox cmbSrvrNm 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   2175
      End
      Begin VB.ComboBox cmbDbNm 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1815
         Width           =   2175
      End
      Begin VB.TextBox TxtPartition 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   4710
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "C"
         Top             =   2100
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox TXT_UCD 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2310
         MaxLength       =   8
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TXT_PWD 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2310
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Frame FRA_KEYS 
         Caption         =   "Frame1"
         Height          =   255
         Left            =   5910
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Frame FRA_NKEYS 
         Caption         =   "Frame1"
         Height          =   165
         Left            =   6000
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TxtCoCd 
         BackColor       =   &H00FFF2DF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2310
         MaxLength       =   3
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
      Begin PicClip.PictureClip PictureClip1 
         Left            =   3720
         Top             =   1020
         _ExtentX        =   1455
         _ExtentY        =   1191
         _Version        =   393216
      End
      Begin MwfCtl.MWCTL_BTN1 CMDOK 
         Default         =   -1  'True
         Height          =   435
         Left            =   900
         TabIndex        =   6
         ToolTipText     =   "Enter Project"
         Top             =   2310
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         ForeColor       =   8388608
         Caption         =   "&OK"
         Pic             =   "EmrFrmJMDI.frx":3D3C1
         ReSizePic       =   0   'False
         BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnDefault      =   -1  'True
      End
      Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
         Height          =   285
         Left            =   6120
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
      End
      Begin MwfCtl.MWCTL_ADC ADC 
         Height          =   495
         Left            =   5160
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         BtnWidth        =   885
         BtnHeight       =   675
         FontSz          =   8.25
      End
      Begin MwfCtl.MWCTL_CTL Ctl 
         Left            =   0
         Top             =   0
         _ExtentX        =   926
         _ExtentY        =   661
      End
      Begin MwfCtl.MWCTL_BTN1 CMDCAN 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   2340
         TabIndex        =   7
         ToolTipText     =   "Exit Project"
         Top             =   2310
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         ForeColor       =   8388608
         Caption         =   "&Cancel"
         Pic             =   "EmrFrmJMDI.frx":3EAD3
         ReSizePic       =   0   'False
         BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnCancel       =   -1  'True
      End
      Begin MSMAPI.MAPIMessages oMapiMsg 
         Left            =   7620
         Top             =   780
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession oMapiSes 
         Left            =   7620
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   720
         TabIndex        =   22
         Top             =   1575
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   21
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label LblPartition 
         BackStyle       =   0  'Transparent
         Caption         =   "Partition (C/F)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   3210
         TabIndex        =   20
         Top             =   2130
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label LblCoCd 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   810
         TabIndex        =   18
         Top             =   345
         Width           =   1485
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   1
         Left            =   810
         TabIndex        =   17
         Top             =   1155
         Width           =   1455
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "User Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   0
         Left            =   810
         TabIndex        =   16
         Top             =   750
         Width           =   1305
      End
      Begin VB.Image ImgDefBtnSet 
         Height          =   1095
         Left            =   2880
         Picture         =   "EmrFrmJMDI.frx":401E5
         Top             =   1920
         Visible         =   0   'False
         Width           =   2385
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4020
      Top             =   30
   End
   Begin ComctlLib.Toolbar TlBar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   8
      Top             =   3105
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton CmdToolBar 
         Height          =   225
         Index           =   0
         Left            =   3330
         TabIndex        =   10
         Top             =   90
         Visible         =   0   'False
         Width           =   500
      End
      Begin MwfCtl.MWCTL_BTN1 CmdCtlToolBar 
         Height          =   375
         Index           =   0
         Left            =   540
         TabIndex        =   9
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         BackColor       =   -2147483633
         DisableColor    =   0
         BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   19
      Top             =   3030
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8123
            Key             =   "ErrMsg"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   265
            MinWidth        =   265
            Key             =   "InMail"
            Object.ToolTipText     =   "No of Unread Mails in Your Inbox"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   35
            Key             =   "ErrPct"
            Object.ToolTipText     =   "Click to View Previous Error"
         EndProperty
      EndProperty
   End
   Begin VB.Menu Parameter 
      Caption         =   "&Parameters"
      Tag             =   "1"
      Visible         =   0   'False
      Begin VB.Menu Param 
         Caption         =   "Parameters"
         Tag             =   "11"
      End
      Begin VB.Menu VarParam 
         Caption         =   "Variable Parameters"
         Tag             =   "12"
      End
      Begin VB.Menu Head 
         Caption         =   "Head Master"
         Tag             =   "13"
      End
      Begin VB.Menu UsrAcss 
         Caption         =   "User Access"
         Tag             =   "14"
      End
      Begin VB.Menu UsrAdm 
         Caption         =   "User Administration"
         Tag             =   "15"
      End
      Begin VB.Menu IdeDfn 
         Caption         =   "IDE Definition"
         Tag             =   "16"
      End
      Begin VB.Menu D1 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MASTERS 
      Caption         =   "&Masters"
      Tag             =   "2"
      Visible         =   0   'False
      Begin VB.Menu Cust 
         Caption         =   "Customer/Supplier"
         Tag             =   "21"
         Visible         =   0   'False
         Begin VB.Menu CustMst 
            Caption         =   "Customer Master"
            Tag             =   "211"
         End
         Begin VB.Menu SuppMst 
            Caption         =   "Supplier Master"
            Tag             =   "212"
         End
         Begin VB.Menu PLMst 
            Caption         =   "Price List Customer Master"
            Tag             =   "213"
         End
         Begin VB.Menu CoMst 
            Caption         =   "Company Master"
            Tag             =   "214"
         End
         Begin VB.Menu D22 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu Rm 
         Caption         =   "Raw Material"
         Tag             =   "22"
         Visible         =   0   'False
         Begin VB.Menu RmMst 
            Caption         =   "Raw Material Master"
            Tag             =   "221"
         End
         Begin VB.Menu RmZMst 
            Caption         =   "Customs RM Master"
            Tag             =   "222"
         End
         Begin VB.Menu Grd 
            Caption         =   "Grade"
            Tag             =   "223"
         End
         Begin VB.Menu VldRmLn 
            Caption         =   "Valid RmLn"
            Tag             =   "224"
         End
         Begin VB.Menu RmIdSz 
            Caption         =   "RmId Sz Desc"
            Tag             =   "225"
         End
         Begin VB.Menu D23 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu Dsg 
         Caption         =   "Design"
         Tag             =   "23"
         Visible         =   0   'False
         Begin VB.Menu DM 
            Caption         =   "Design Master"
            Tag             =   "231"
         End
         Begin VB.Menu SM 
            Caption         =   "Design Sketch Master"
            Tag             =   "232"
         End
         Begin VB.Menu PM 
            Caption         =   "Design Component Master"
            Tag             =   "233"
         End
         Begin VB.Menu DsgPrm 
            Caption         =   "Design Parameters"
            Tag             =   "234"
         End
         Begin VB.Menu DsgPrmCat 
            Caption         =   "Design Catalogue"
            Tag             =   "235"
         End
         Begin VB.Menu DsgView 
            Caption         =   "Detailed Design View"
            Tag             =   "236"
         End
         Begin VB.Menu D24 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RtMst 
         Caption         =   "Rate Master"
         Tag             =   "24"
         Visible         =   0   'False
         Begin VB.Menu RmRt 
            Caption         =   "Raw Material Rate Master"
            Tag             =   "241"
         End
         Begin VB.Menu ILC 
            Caption         =   "Increment Logic Rate Chart"
            Tag             =   "244"
         End
         Begin VB.Menu RmRtHist 
            Caption         =   "RM Rate History"
            Tag             =   "242"
         End
         Begin VB.Menu LabRt 
            Caption         =   "Labour Rate Master"
            Tag             =   "243"
         End
         Begin VB.Menu D25 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu SubMst 
         Caption         =   "Sub Masters"
         Tag             =   "25"
         Begin VB.Menu Loc 
            Caption         =   "Location Master"
            Tag             =   "251"
         End
         Begin VB.Menu DtTbl 
            Caption         =   "Date Table"
            Tag             =   "252"
         End
         Begin VB.Menu PrdPts 
            Caption         =   "Production Points Master"
            Tag             =   "253"
         End
         Begin VB.Menu LossEntry 
            Caption         =   "Loss Entry"
            Tag             =   "254"
         End
         Begin VB.Menu Bnk 
            Caption         =   "Bank/Agent Master"
            Tag             =   "255"
         End
         Begin VB.Menu D26 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu LsRec 
         Caption         =   "Loss Recovery"
         Tag             =   "26"
      End
      Begin VB.Menu D2 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu JTTCS 
      Caption         =   "&Transactions"
      Tag             =   "8"
      Visible         =   0   'False
      Begin VB.Menu Quotation 
         Caption         =   "Quotation/PDCM"
         Tag             =   "87"
         Begin VB.Menu SPL 
            Caption         =   "Supplier PDCM"
            Tag             =   "871"
         End
         Begin VB.Menu SQT 
            Caption         =   "Suplier Quotation"
            Tag             =   "872"
         End
         Begin VB.Menu PrcLst 
            Caption         =   "Customer PDCM"
            Tag             =   "322"
         End
         Begin VB.Menu Quot 
            Caption         =   "Customer Quotation"
            Tag             =   "874"
         End
         Begin VB.Menu JQB 
            Caption         =   "Design Stock Quotation"
            Tag             =   "875"
         End
         Begin VB.Menu D987 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu JS 
         Caption         =   "Sales"
         Tag             =   "81"
         Begin VB.Menu JSO 
            Caption         =   "Sales Order"
            Tag             =   "811"
         End
         Begin VB.Menu JSA 
            Caption         =   "Sales"
            Tag             =   "812"
         End
         Begin VB.Menu JMI 
            Caption         =   "Memo Issue"
            Tag             =   "813"
         End
         Begin VB.Menu JMR 
            Caption         =   "Memo Return"
            Tag             =   "814"
         End
         Begin VB.Menu JMS 
            Caption         =   "Memo Sales"
            Tag             =   "815"
         End
         Begin VB.Menu JST 
            Caption         =   "Sales Return"
            Tag             =   "816"
         End
         Begin VB.Menu D31 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu JP 
         Caption         =   "Purchase"
         Tag             =   "82"
         Visible         =   0   'False
         Begin VB.Menu JPO 
            Caption         =   "Purchase Order"
            Tag             =   "821"
         End
         Begin VB.Menu JPR 
            Caption         =   "Purchase"
            Tag             =   "822"
         End
         Begin VB.Menu JCI 
            Caption         =   "Consignment Inward"
            Tag             =   "823"
         End
         Begin VB.Menu JCR 
            Caption         =   "Consignment Return"
            Tag             =   "824"
         End
         Begin VB.Menu JCP 
            Caption         =   "Consignment Purchase"
            Tag             =   "825"
         End
         Begin VB.Menu JPT 
            Caption         =   "Purchase Returns"
            Tag             =   "826"
         End
         Begin VB.Menu D32 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu JMV 
         Caption         =   "Design Movement"
         Tag             =   "83"
      End
      Begin VB.Menu JCN 
         Caption         =   "Design Conversion"
         Tag             =   "84"
      End
      Begin VB.Menu JMT 
         Caption         =   "Design Melting"
         Tag             =   "85"
      End
      Begin VB.Menu InvHdJT 
         Caption         =   "InvHd"
         Tag             =   "89"
      End
      Begin VB.Menu MnuJTINOUT 
         Caption         =   "&Transfer"
         Tag             =   "86"
         Begin VB.Menu JIN 
            Caption         =   "Transfer In"
            Tag             =   "861"
         End
         Begin VB.Menu JOT 
            Caption         =   "Transfer Out"
            Tag             =   "863"
         End
         Begin VB.Menu D864 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu Rpr 
         Caption         =   "Repair"
         Tag             =   "88"
         Begin VB.Menu JRI 
            Caption         =   "Repair In"
            Tag             =   "881"
         End
         Begin VB.Menu JRO 
            Caption         =   "Repair Out"
            Tag             =   "882"
         End
         Begin VB.Menu dffg 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu D3 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu JL 
      Caption         =   "Loose Transactions"
      Tag             =   "9"
      Visible         =   0   'False
      Begin VB.Menu LP 
         Caption         =   "Purchase"
         Tag             =   "91"
         Begin VB.Menu LPO 
            Caption         =   "Loose Purchase Order"
            Tag             =   "911"
         End
         Begin VB.Menu LPR 
            Caption         =   "Loose Purchase"
            Tag             =   "912"
         End
         Begin VB.Menu LPT 
            Caption         =   "Loose Purchase Return"
            Tag             =   "913"
         End
         Begin VB.Menu DUM914 
            Caption         =   "DUM"
            Tag             =   "914"
         End
      End
      Begin VB.Menu LS 
         Caption         =   "Sales"
         Tag             =   "92"
         Begin VB.Menu LSO 
            Caption         =   "Loose Sales Order"
            Tag             =   "921"
         End
         Begin VB.Menu LSA 
            Caption         =   "Loose Sales"
            Tag             =   "922"
         End
         Begin VB.Menu LST 
            Caption         =   "Loose Sales Return"
            Tag             =   "923"
         End
         Begin VB.Menu DUM924 
            Caption         =   "DUM"
            Tag             =   "924"
         End
      End
      Begin VB.Menu LM 
         Caption         =   "Loose Memo"
         Tag             =   "93"
         Begin VB.Menu LMS 
            Caption         =   "Loose Memo Sales"
            Tag             =   "932"
         End
         Begin VB.Menu LMI 
            Caption         =   "Loose Memo Issue"
            Tag             =   "931"
         End
         Begin VB.Menu LMR 
            Caption         =   "Loose Memo Return"
            Tag             =   "933"
         End
         Begin VB.Menu DUM934 
            Caption         =   "DUM"
            Tag             =   "934"
         End
      End
      Begin VB.Menu LC 
         Caption         =   "Consignment"
         Tag             =   "94"
         Begin VB.Menu LCI 
            Caption         =   "Loose Consignment Inward"
            Tag             =   "941"
         End
         Begin VB.Menu LCP 
            Caption         =   "Loose Consignment Purchase"
            Tag             =   "942"
         End
         Begin VB.Menu LCR 
            Caption         =   "Loose Consignment Return"
            Tag             =   "943"
         End
         Begin VB.Menu DUM944 
            Caption         =   "DUM"
            Tag             =   "944"
         End
      End
      Begin VB.Menu LT 
         Caption         =   "Transfer"
         Tag             =   "95"
         Begin VB.Menu LIN 
            Caption         =   "Transfer In"
            Tag             =   "951"
         End
         Begin VB.Menu LOT 
            Caption         =   "Transfer Out"
            Tag             =   "952"
         End
         Begin VB.Menu D953 
            Caption         =   "DUM"
            Tag             =   "953"
         End
      End
      Begin VB.Menu LMV 
         Caption         =   "Loose Movement"
         Tag             =   "96"
      End
      Begin VB.Menu LCN 
         Caption         =   "Loose Conversion"
         Tag             =   "97"
      End
      Begin VB.Menu D96 
         Caption         =   "DUM"
         Tag             =   "99"
      End
   End
   Begin VB.Menu Rep 
      Caption         =   "Reports"
      Tag             =   "4"
      Visible         =   0   'False
      Begin VB.Menu RepJtTxnLstRpt 
         Caption         =   "Transaction Listing (Jwlry)"
         Tag             =   "J41"
      End
      Begin VB.Menu CstRpt 
         Caption         =   "Costing Reports (Jwlry)"
         Tag             =   "J42"
         Begin VB.Menu RepJtDet 
            Caption         =   "Detail Costing"
            Tag             =   "J421"
         End
         Begin VB.Menu RepJtAvg 
            Caption         =   "Average Costing"
            Tag             =   "J422"
         End
         Begin VB.Menu RepJTCat 
            Caption         =   "Catalogue format"
            Tag             =   "J423"
         End
         Begin VB.Menu jPrfMrgn 
            Caption         =   "Profit Margin Report"
            Tag             =   "J424"
         End
         Begin VB.Menu D412 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu PndVchrJwlry 
         Caption         =   "Pending Voucher (Jwlry)"
         Tag             =   "J43"
         Begin VB.Menu RepJtPnd 
            Caption         =   "Pending Voucher"
            Tag             =   "J431"
         End
         Begin VB.Menu RepJtPndRm 
            Caption         =   "Pending Voucher Rm (Rm Level)"
            Tag             =   "J432"
         End
         Begin VB.Menu RepJtSOPndRm 
            Caption         =   "Pending Voucher Rm (Dsg Level)"
            Tag             =   "J434"
         End
         Begin VB.Menu RepJtOpnOrd 
            Caption         =   "Open Reference Vchrs"
            Tag             =   "J433"
         End
         Begin VB.Menu D413 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu StkRptJwlry 
         Caption         =   "Stock Reports (Jwlry)"
         Tag             =   "J44"
         Begin VB.Menu RepJTStkLed 
            Caption         =   "Jewellery Stock Ledger"
            Tag             =   "J441"
         End
         Begin VB.Menu RepJtStkRm 
            Caption         =   "Jwellery Stock RM"
            Tag             =   "J442"
         End
         Begin VB.Menu RepJtDsgFlashStk 
            Caption         =   "Design Flash Stock"
            Tag             =   "J443"
         End
         Begin VB.Menu RepJtNetStk 
            Caption         =   "Net Jwlry Stock Requirement"
            Tag             =   "J444"
         End
         Begin VB.Menu D414 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu AnaRptJwlry 
         Caption         =   "Analysis Reports (Jwlry)"
         Tag             =   "J45"
         Begin VB.Menu RepJtVchAnaDsg 
            Caption         =   "Voucher Analysis - Design Wise"
            Tag             =   "J451"
         End
         Begin VB.Menu RepJtVchAnaRm 
            Caption         =   "Voucher Analysis - RM"
            Tag             =   "J452"
         End
         Begin VB.Menu RepJtVchAnaLab 
            Caption         =   "Voucher Analysis - Labour"
            Tag             =   "J453"
         End
         Begin VB.Menu D415 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepJtLclInv 
         Caption         =   "Invoice Printing (Jwlry)"
         Tag             =   "J46"
      End
      Begin VB.Menu ZoomRpt 
         Caption         =   "Zoom Reports (Jwlry)"
         Tag             =   "J47"
         Begin VB.Menu RepJtZoom 
            Caption         =   "Jewellery Trading Zoom"
            Tag             =   "J471"
         End
         Begin VB.Menu RepJtStkZoom 
            Caption         =   "Jewellery Trading Stock Zoom"
            Tag             =   "J472"
         End
         Begin VB.Menu D416 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu JtStkReco 
         Caption         =   "Stock Reco (Jwlry)"
         Tag             =   "J48"
         Begin VB.Menu JtStkRecoEntry 
            Caption         =   "Stock Reco Entry"
            Tag             =   "J481"
         End
         Begin VB.Menu RepJTStkReco 
            Caption         =   "Stock Reco Report"
            Tag             =   "J482"
         End
         Begin VB.Menu D43 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu REPLRM 
         Caption         =   "Loose Trading Reports"
         Tag             =   "J49"
         Begin VB.Menu VchrReg 
            Caption         =   "Voucher - Register"
            Tag             =   "J491"
            Begin VB.Menu RLRMVPR 
               Caption         =   "Voucher Printing"
               Tag             =   "J4911"
            End
            Begin VB.Menu RLRMTL 
               Caption         =   "Transaction Listing - Register"
               Tag             =   "J4912"
            End
            Begin VB.Menu D421 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu PndVchrLoose 
            Caption         =   "Pending Voucher (Loose)"
            Tag             =   "J492"
            Begin VB.Menu RLRMPO 
               Caption         =   "Pending Voucher (Loose)"
               Tag             =   "J4921"
            End
            Begin VB.Menu D422 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu StkRptLoose 
            Caption         =   "Stock Reports (Loose)"
            Tag             =   "J493"
            Begin VB.Menu RLRMSTK 
               Caption         =   "Stock Ledger"
               Tag             =   "J4931"
            End
            Begin VB.Menu D423 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu DRLRM 
            Caption         =   "DUM"
            Tag             =   "476"
         End
      End
      Begin VB.Menu RepJtTagPrn 
         Caption         =   "Tag Printing (Jwlry)"
         Tag             =   "J4A"
      End
      Begin VB.Menu D4 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu RepLst 
      Caption         =   "A&ddOn Menus"
      Tag             =   "5"
      Visible         =   0   'False
   End
   Begin VB.Menu Mail 
      Caption         =   "Mai&l"
      Tag             =   "6"
      Visible         =   0   'False
      Begin VB.Menu UsrMail 
         Caption         =   "User Mail"
         Tag             =   "61"
         Visible         =   0   'False
      End
      Begin VB.Menu MailSignIn 
         Caption         =   "Sign In"
         Tag             =   "62"
         Visible         =   0   'False
      End
      Begin VB.Menu SndMail 
         Caption         =   "Send Mail"
         Tag             =   "63"
         Visible         =   0   'False
      End
      Begin VB.Menu MailSignOut 
         Caption         =   "Sign Out"
         Tag             =   "64"
         Visible         =   0   'False
      End
      Begin VB.Menu D6 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu MnuTlBar 
      Caption         =   "T&oolBar"
      Tag             =   "7"
      Visible         =   0   'False
   End
   Begin VB.Menu WINDOW 
      Caption         =   "&Window"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu CSC 
         Caption         =   "Cascade"
      End
      Begin VB.Menu TIL 
         Caption         =   "Tile"
      End
      Begin VB.Menu GENMNU 
         Caption         =   "Generate Menu"
      End
      Begin VB.Menu TRACK 
         Caption         =   "Track Status"
         Tag             =   "6793"
      End
      Begin VB.Menu D8 
         Caption         =   "DUM"
      End
   End
   Begin VB.Menu QUIT 
      Caption         =   "&Quit"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "EmrFrmJMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ms_MnuForm  As String
 Dim mo_MstConn As MwfLib.MDOConnection
 
'*** Jenny Color
Dim mo_MDIPic As Picture
Dim mo_CsTlBarPic As Picture
Dim mo_TlBarColl As Collection   '*** (Bef 12/08/05)

'Public mo_TlBarColl As Collection '*** (12/08/05)
'*** Jenny Color
Public Function TlBarColl() As Collection
    
   Set TlBarColl = mo_TlBarColl
End Function

Private Sub CmdCan_Click()
  Unload Me
End Sub

'VK.27-New menu "Pending voucher rm(Dsg level)" added.
Private Sub RepJtSOPndRm_Click()
Call FormSelect("RepJtSOPndRm")
End Sub

Private Sub RepJtStkZoom_Click()
Call FormSelect("RepJtStkZoom")

End Sub

Private Sub cmbSrvrNm_Click()
Call FillDbCmb      '3.11.4
End Sub

'Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As ComctlLib.ButtonMenu)
'  '*** Jen ***
'  'Place a toolBar on a cool bar. Add Buttons to the ToolBar. Add ButtonMenus to the button.
'  'If need be add Popup Menus also.
'  'In the general tab of the ToolBar properties set the TextAlignment as '1 - tbrTextAlignRight'
'  'In the buttons tab, keep the Style of the button as '5 - tbrDropdown'
'  'If a popup menu has to be displayed for a buttonmenu then a key has to be specified
'  'for the buttonmenu.
'
'  If ButtonMenu.Key = "A" Then PopupMenu Cust
'  '*** Jen ***
'End Sub

Private Sub BM_Click()
  ' **** Zubin 211 **** '
  Call FormSelect("BM")
  ' **** Zubin 211 **** '
End Sub

'*** Jay 3.4[DP]   '*** (Jen 3.01)
Private Sub DsgView_Click()
  Call FormSelect("DsgView")
End Sub
'*** Jay 3.4[DP]    '*** (Jen 3.01)

Private Sub DsgPrm_Click()
'*** Jay 3.4[DP]
  Call FormSelect("DsgPrm")
End Sub

Private Sub Jin_Click()
    Call FormSelect("JIN")
End Sub
Private Sub Jot_Click()
  '*** Jay 3.3(FG) ***
  Call FormSelect("JOT")
End Sub

Private Sub InvHdJT_Click()
 ' Call FormSelect("InvHdJT")
End Sub
Private Sub DsgPrmCat_Click()
'*** Jay 3.4[DP]
  Call FormSelect("DsgPrmCat")
End Sub
Private Sub ClrDfn_Click()
  Call FormSelect("CLRDFN")
End Sub
Private Sub ClrSel_Click()
  Call FormSelect("CLRSEL")
End Sub
Private Sub FM_Click()
  Call FormSelect("FM")
End Sub
Private Sub Grd_Click()
  Call FormSelect("Grd")
End Sub
Private Sub Head_Click()
  Call FormSelect("Head")
End Sub
Private Sub CustMst_Click()
  Call FormSelect("CustMst")
End Sub
Private Sub CtlgMst_Click()
  Call FormSelect("CtlgMst")
End Sub
' Zubin 213 (Shifted to Lising Reports)
'Private Sub RC_Click()
'  Call FormSelect("RC")
'End Sub
' Zubin 213 (Shifted to Lising Reports)
'Private Sub RepBagReco_Click()
'  Call FormSelect("RepBagReco")
'End Sub
Private Sub ILC_Click()
  '*** Jay 3.1.2 ***
  Call FormSelect("ILC")
End Sub

Private Sub RepBvRm_Click()
  ' **** Zubin 2.11 **** '
  Call FormSelect("RepBvRm")
' **** Zubin 2.11 **** '
End Sub

Private Sub JRI_Click()
Call FormSelect("JRI")
End Sub

Private Sub JRO_Click()
Call FormSelect("JRO")
End Sub
'pg.19
Private Sub JtStkRecoEntry_Click()
Call FormSelect("JtStkRecoEntry")
End Sub

Private Sub LIN_Click()
Call FormSelect("LIN")
End Sub
Private Sub LOT_Click()
Call FormSelect("LOT")
End Sub
Private Sub LMI_Click()
Call FormSelect("LMI")
End Sub
Private Sub LMR_Click()
Call FormSelect("LMR")
End Sub
Private Sub LMS_Click()
Call FormSelect("LMS")
End Sub
Private Sub LPO_Click()
Call FormSelect("LPO")
End Sub
Private Sub LPR_Click()
Call FormSelect("LPR")
End Sub
Private Sub LPT_Click()
Call FormSelect("LPT")
End Sub
Private Sub LSO_Click()
Call FormSelect("LSO")
End Sub
Private Sub LSA_Click()
Call FormSelect("LSA")
End Sub
Private Sub LST_Click()
Call FormSelect("LST")
End Sub
Private Sub RepFgBagLed_Click()
  ' Zubin 212 (Option Moved From Report Listing)
  Call FormSelect("RepFgBagLed")
End Sub

Private Sub RepFltBagPrn_Click()
' **** Zubin 211 **** '
  Call FormSelect("RepFltBagPrn")
' **** Zubin 211 **** '
End Sub

Private Sub JCN_Click()
  Call FormSelect("JCN")
End Sub
Private Sub LCN_Click()
  Call FormSelect("LCN")
End Sub
Private Sub LMV_Click()
  Call FormSelect("LMV")
End Sub
Private Sub RepLclInv_Click()
  Call FormSelect("RepLclInv")
End Sub

Private Sub JCI_Click()
   Call FormSelect("JCI")
End Sub
Private Sub LCI_Click()
   Call FormSelect("LCI")
End Sub
Private Sub LCR_Click()
   Call FormSelect("LCR")
End Sub
Private Sub LCP_Click()
   Call FormSelect("LCP")
End Sub
Private Sub JCP_Click()
Call FormSelect("JCP")
End Sub

Private Sub JCR_Click()
Call FormSelect("JCR")
End Sub

Private Sub JMI_Click()
Call FormSelect("JMI")
End Sub

Private Sub JMR_Click()
Call FormSelect("JMR")
End Sub

Private Sub JMS_Click()
Call FormSelect("JMS")
End Sub

Private Sub JMV_Click()
Call FormSelect("JMV")
End Sub

Private Sub JMT_Click()
Call FormSelect("JMT")
End Sub


Private Sub JPO_Click()
Call FormSelect("JPO")
End Sub

Private Sub JPR_Click()
Call FormSelect("JPR")
End Sub

Private Sub JPT_Click()
Call FormSelect("JPT")
End Sub

Private Sub JSA_Click()
Call FormSelect("JSA")
End Sub

Private Sub JSO_Click()
Call FormSelect("JSO")
End Sub

Private Sub JST_Click()
Call FormSelect("JST")
End Sub
'vk.22 JQB added
Private Sub JQB_Click()
Call FormSelect("JQB")
End Sub

Private Sub RepJtAvg_Click()
Call FormSelect("RepJtAvg")
End Sub

Private Sub RepJTCat_Click()
Call FormSelect("RepJtCat")
End Sub

Private Sub RepJtDet_Click()
Call FormSelect("RepJtDet")
End Sub

Private Sub RepJtGrBk_Click()
Call FormSelect("RepJtGrBk")
End Sub

Private Sub RepJtGrFrnt_Click()
Call FormSelect("RepJtGrFrnt")
End Sub

Private Sub RepJtGsp_Click()
Call FormSelect("RepJtGsp")
End Sub

Private Sub RepJtInv_Click()
Call FormSelect("RepJtInv")
End Sub

Private Sub RepJtLclInv_Click()
Call FormSelect("RepJtLclInv")
End Sub

Private Sub RepJtNetStk_Click()
Call FormSelect("RepJtNetStk")
End Sub

Private Sub RepJtOpnOrd_Click()
Call FormSelect("RepJtOpnOrd")
End Sub

Private Sub RepJtPckLst_Click()
Call FormSelect("RepJtPckLst")
End Sub

Private Sub RepJtPnd_Click()
Call FormSelect("RepJtPnd")
End Sub

Private Sub RepJtPndRm_Click()
Call FormSelect("RepJtPndRm")
End Sub

Private Sub RepJtShpBill_Click()
Call FormSelect("RepJtShpBill")
End Sub

Private Sub RepJTStkLed_Click()
Call FormSelect("RepJTStkLed")
End Sub

Private Sub RepJtStkRm_Click()
Call FormSelect("RepJTStkRm")
End Sub

'VK.25 - New menu added
Private Sub RepJtTagPrn_Click()
Call FormSelect("RepJtTagPrn")
End Sub

Private Sub RepJtVchAnaDsg_Click()
Call FormSelect("RepJTVchAnaDsg")
End Sub

Private Sub RepJtVchAnaLab_Click()
Call FormSelect("RepJTVchAnaLab")
End Sub

Private Sub RepJtZoom_Click()
Call FormSelect("RepJtZoom")
End Sub


Private Sub RepJtVchAnaRm_Click()
Call FormSelect("RepJTVchAnaRm")
End Sub

Private Sub RLRMPO_Click()
Call FormSelect("RLRMPO")
End Sub

Private Sub RLRMSTK_Click()
Call FormSelect("RLRMSTK")
End Sub

Private Sub RLRMTL_Click()
Call FormSelect("RLRMTL")
End Sub

Private Sub RLRMVPR_Click()
Call FormSelect("RLRMVPR")
End Sub

Private Sub SPL_Click()
Call FormSelect("SPL")
End Sub

Private Sub Quot_Click()
Call FormSelect("Quot")
End Sub
Private Sub SQT_Click()
Call FormSelect("SQT")
End Sub
Private Sub RepJTStkReco_Click()
Call FormSelect("RepJTStkReco")
End Sub
Private Sub RepLst_Click()
' ########################  Manoj #### Ver: 2.0.6 #### Date: 02/02/2004  ########################
  ' Form with Menu Placed for all Listing Reports
  Call FormSelect("RepLst")
' ###########################################  Manoj  ###########################################
End Sub
Private Sub RepRmStkReq_Click()
' Include later
  Call FormSelect("RepRmStkReq")
End Sub
Private Sub REPNETSTK_Click()
  Call FormSelect("RepNetStk")
End Sub
Private Sub RepTagPrnBagFg_Click()
  Call FormSelect("RepTagPrnBagFg")
End Sub
Private Sub RepTagPrnFg_Click()
  Call FormSelect("RepTagPrnFg")
End Sub

' Zubin 250308 (Emr 3.02.01)
Private Sub RepTravExcel_Click()
  Call FormSelect("RepTravExcel")
End Sub

Private Sub RmRtHist_Click()
  '*************************** ZUBIN **************************
  ' 27th Nov 2003, EMR206 - Menu for RM Rate History
  Call FormSelect("RmRtHist")
  '*************************** ZUBIN **************************
End Sub
'********* Urmi Ech Diamond Sales ***********'
Private Sub Sa_Click()
  Call FormSelect("SA")
End Sub
Private Sub SalOrdAck_Click()
  ' Zubin 212P2
  Call FormSelect("SalOrdAck")
End Sub

'********************************************'
Private Sub SuppMst_Click()
  Call FormSelect("SuppMst")
End Sub
Private Sub PLMst_Click()
  Call FormSelect("PLMst")
End Sub
Private Sub CoMst_Click()
  Call FormSelect("CoMst")
End Sub
Private Sub ContractorMst_Click()
  Call FormSelect("ContractorMst")
End Sub
Private Sub Param_Click()
  Call FormSelect("Param")
End Sub
Private Sub Tree_Click()
  Call FormSelect("Tree")
End Sub

Private Sub TXT_UCD_LostFocus()
    '3.11.4
    Dim wSrvrNm As String, i As Integer
'    Dim wo_MstConn As MwfLib.MDOConnection
'    Set wo_MstConn = MWLib.GetMwMstConn(ctProjectName)
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
    End If
    wSrvrNm = mo_MstConn.GetFldVal("Select Top 1 MuSrvrNm From MUsrMst Where MuSysId='" + ctProjectName + "' And MuUsrId In ('','" + TXT_UCD + "') Order by MuUsrId Desc ")
    For i = 0 To cmbSrvrNm.ListCount - 1
            If UCase(cmbSrvrNm.List(i)) = UCase(wSrvrNm) Then
                cmbSrvrNm.ListIndex = i
                Call FillDbCmb
                Exit For
            End If
    Next
    'Set wo_MstConn = Nothing
End Sub

Private Sub TxtPartition_KeyPress(KeyAscii As Integer)
    '****** Sachin 3.02 [26/11/07] - Speed Optimization
    KeyAscii = Asc(UCase$(Chr$((KeyAscii))))
End Sub
'Private Sub MicParam_Click()
'  Dim wfrm As New EmrFrmParam
'  ShowForm wfrm, "MicParam"
'  wfrm.Caption = "Microway Parameters"
'End Sub
'Private Sub SysParam_Click()
'  Dim wfrm As New EmrFrmParam
'  ShowForm wfrm, "SysParam"
'  wfrm.Caption = "System Parameters"
'End Sub
'Private Sub OthParam_Click()
'  Dim wfrm As New EmrFrmParam
'  ShowForm wfrm, "OthParam"
'  wfrm.Caption = "Other Parameters"
'End Sub
Private Sub VarParam_Click()
  Call FormSelect("VarParam")
End Sub
'Private Sub VarMicParam_Click()
'  Dim wfrm As New EmrFrmvParam
'  ShowForm wfrm, "VarMicParam"
'  wfrm.Caption = "Variable Microway Parameters"
'End Sub
'Private Sub VarSysParam_Click()
'  Dim wfrm As New EmrFrmvParam
'  ShowForm wfrm, "VarSysParam"
'  wfrm.Caption = "Variable System Parameters"
'End Sub
'Private Sub VarOthParam_Click()
'  Dim wfrm As New EmrFrmvParam
'  ShowForm wfrm, "VarOthParam"
'  wfrm.Caption = "Variable Other Parameters"
'End Sub
Private Sub UsrAcss_Click()
  Call FormSelect("UsrAcss")
End Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 17/02/2004  ########################
Private Sub UsrAdm_Click()
  Call FormSelect("UsrAdm")
End Sub
' ###########################################  Manoj  ###########################################

'*** (Jen 3.01)
Private Sub IdeDFN_Click()
  Call FormSelect("IdeDfn")
End Sub
Private Sub IdeUsrRight_Click()
  Call FormSelect("IdeUsrRight")
End Sub
'*** (Jen 3.01)

Private Sub RmMst_Click()
  Call FormSelect("RmMst")
End Sub
Private Sub RmZMst_Click()
  Call FormSelect("RmZMst")
End Sub
Private Sub DM_Click()
  Call FormSelect("DM")
End Sub
Private Sub PM_Click()
  Call FormSelect("PM")
End Sub
Private Sub SM_Click()
  Call FormSelect("SM")
End Sub
Private Sub RmRt_Click()
  Call FormSelect("RmRt")
End Sub
Private Sub LabRt_Click()
  Call FormSelect("LabRt")
End Sub
Private Sub Loc_Click()
  Call FormSelect("Loc")
End Sub
Private Sub DtTbl_Click()
  Call FormSelect("DtTbl")
End Sub
Private Sub Bnk_Click()
  Call FormSelect("Bnk")
End Sub
'Private Sub MnthRecPrc_Click()
'  Dim wfrm As New EmrFrmMonthlyRecPrc
'  ShowForm wfrm, "MnthRecPrc"
'  wfrm.Caption = "Monthly Recovery %"
'End Sub
Private Sub Bag_Click()
  Call FormSelect("Bag")
End Sub
Private Sub SalOrd_Click()
  Call FormSelect("SalOrd")
End Sub
Private Sub PrcLst_Click()
  Call FormSelect("PrcLst")
End Sub
'Private Sub Quot_Click()
'  Call FormSelect("Quot")
'End Sub
Private Sub QB_Click()
  Call FormSelect("QB")
End Sub
Private Sub SkQuot_Click()
  Call FormSelect("SkQuot")
End Sub
Private Sub RetMem_Click()
  Call FormSelect("RetMem")
End Sub
Private Sub ImpData_Click()
  ' Dim wfrm As New FrmDataTransfer
  ' Dim wfrm As New EmrFrmImpData
  'Dim wfrm As New EmrFrmImpDataGlb
  'Dim wfrm As New EmrFrmImpDataFine
  'Dim wfrm As New EmrFrmImpDataFineOne
  'Dim wfrm As New EmrFrmOrdZoom
  'Dim wfrm As New EchFrpEchTag

  'Dim wfrm As New EmrFrmPpcOrdPln
  'Dim wfrm As New EmrFrpDsgCat
  '*** Manoj
  'Dim wfrm As New EmrFrmTagFmt
  '*** Manoj

  'ShowForm wfrm, "ImpData"
End Sub

' ***** Manali 3.10.0 - 10/05/12 - Reset Autogen shifted to AddonMnus
''Private Sub ResAuto_Click()
''  Call FormSelect("ResAuto")
''End Sub

Private Sub CpyRt_Click()
  Call FormSelect("CpyRt")
End Sub
Private Sub DT_Click()
  Call FormSelect("DT")
End Sub
Private Sub BV_Click()
  Call FormSelect("BV")
End Sub
Private Sub BA_Click()
  Call FormSelect("BA")
End Sub
Private Sub BS_Click()
  Call FormSelect("BS")
End Sub
Private Sub REJ_Click()
  Call FormSelect("REJ")
End Sub
Private Sub MLT_Click()
  Call FormSelect("MLT")
End Sub
Private Sub PrdPts_Click()
  Call FormSelect("PrdPts")
End Sub
Private Sub LossEntry_Click()
  Call FormSelect("LossEntry")
End Sub

'Private Sub SP_Click()
'  Dim wfrm As New EmrFrmTxn
'  wfrm.Caption = "Split Movement"
'  ShowForm wfrm, "SP"
'  wfrm.WindowState = 2
'End Sub

Private Sub PR_Click()
  Call FormSelect("PR")
End Sub
Private Sub Tz_Click()
  Call FormSelect("Tz")
End Sub
Private Sub MV_Click()
  Call FormSelect("MV")
End Sub
'Private Sub REC_Click()
'  Dim wfrm As New EmrFrmTxnRm
'  wfrm.Caption = "Recovery"
'  ShowForm wfrm, "REC"
'  'wfrm.WindowState = 2
'End Sub
Private Sub CNV_Click()
  Call FormSelect("CNV")
End Sub
Private Sub FB_Click()
  Call FormSelect("FB")
End Sub

'Private Sub FR_Click()
'  Dim wfrm As New EmrFrmFg
'  wfrm.Caption = "Finished Goods (RM)"
'  ShowForm wfrm, "FR"
'  'wfrm.WindowState = 2
'End Sub

Private Sub FWB_Click()
  Call FormSelect("FWB")
End Sub
'Private Sub FWR_Click()
'  Dim wfrm As New EmrFrmFg
'  wfrm.Caption = "Finished Goods (RM Return)"
'  ShowForm wfrm, "FWR"
'  'wfrm.WindowState = 2
'End Sub
Private Sub RepInvCustStmt_Click()
'*************************** Geeta **************************
' ***Emr206****************Menu for Customer Statement*******
  Call FormSelect("RepInvCustStmt")
'***************Geeta*****************
End Sub

Private Sub RepTreePerform_Click()
  '***************Geeta*****************
  Call FormSelect("RepTreePerform")
End Sub
Private Sub RepTreeInvst_Click()
  Call FormSelect("RepTreeInvst")
  '***************Geeta*****************
End Sub
Private Sub In_Click()
  Call FormSelect("IN")
End Sub
Private Sub InAllBags_Click()
  Call FormSelect("InAllBags")
End Sub
Private Sub InDsg_Click()
  Call FormSelect("InDsg")
End Sub

Private Sub InCustmMtch_Click()
  Call FormSelect("InCustmMtch")
End Sub
Private Sub CTBIn_Click()
  Call FormSelect("CTBIn")
End Sub
Private Sub CTBInAllBags_Click()
  Call FormSelect("CTBInAllBags")
End Sub
Private Sub CTBInDsg_Click()
  Call FormSelect("CTBInDsg")
End Sub
Private Sub RepPrm_Click()
  Call FormSelect("RepPrm")
End Sub
'Private Sub RepMicPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "Microway Parameter Listing"
'  ShowForm wfrm, "RepMicPrm"
'End Sub
'Private Sub RepSysPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "System Parameter Listing"
'  ShowForm wfrm, "RepSysPrm"
'End Sub
'Private Sub RepOthPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "Other Parameter Listing"
'  ShowForm wfrm, "RepOthPrm"
'End Sub
Private Sub RepVarPrm_Click()
  Call FormSelect("RepVarPrm")
End Sub
'Private Sub RepVarMicPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "Variable Microway Parameter Listing"
'  ShowForm wfrm, "RepVarMicPrm"
'End Sub
'Private Sub RepVarSysPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "Variable System Parameter Listing"
'  ShowForm wfrm, "RepVarSysPrm"
'End Sub
'Private Sub RepVarOthPrm_Click()
'  Dim wfrm As New EmrFrpParam
'  wfrm.Caption = "Variable Other Parameter Listing"
'  ShowForm wfrm, "RepVarOthPrm"
'End Sub
Private Sub RepCustLst_Click()
  Call FormSelect("RepCustLst")
End Sub
Private Sub RepUsrAccess_Click()
'******Geeta****************
  Call FormSelect("RepUsrAccess")
End Sub

'Private Sub RepSuppLst_Click()
'  Dim wfrm As New EmrFrpCustMst
'  wfrm.Caption = "Supplier Listing"
'  ShowForm wfrm, "RepSuppLst"
'End Sub
'Private Sub RepPLLst_Click()
'  Dim wfrm As New EmrFrpCustMst
'  wfrm.Caption = "Price List Customer Listing"
'  ShowForm wfrm, "RepPLLst"
'End Sub
'Private Sub RepCoLst_Click()
'  Dim wfrm As New EmrFrpCustMst
'  wfrm.Caption = "Company Listing"
'  ShowForm wfrm, "RepCoLst"
'End Sub
Private Sub RepCustDsg_Click()
  Call FormSelect("RepCustDsg")
End Sub
Private Sub RepCustAly_Click()
  Call FormSelect("RepCustAly")
End Sub
Private Sub RepRmMst_Click()
  Call FormSelect("RepRmMst")
End Sub
'Private Sub RepRm_Click()
'  Dim wfrm As New EmrFrpRmMst
'  wfrm.Caption = "Raw Material/ Custom Master Listing"
'  ShowForm wfrm, "RepRm"
'End Sub
'Private Sub RepRmZ_Click()
'  Dim wfrm As New EmrFrpRmMst
'  wfrm.Caption = "Custom Code Listing"
'  ShowForm wfrm, "RepRmZ"
'End Sub

Private Sub RepDsg_Click()
  Call FormSelect("RepDsg")
End Sub
Private Sub RepDsgCrd_Click()
  Call FormSelect("RepDsgCrd")
End Sub
Private Sub RepDsgCat_Click()
  Call FormSelect("RepDsgCat")
End Sub


'Private Sub RepDM_Click()
'  Dim wfrm As New EmrFrpDsgLst
'  wfrm.Caption = "Design Master Listing"
'  ShowForm wfrm, "RepDM"
'End Sub
'Private Sub RepSM_Click()
'  Dim wfrm As New EmrFrpDsgLst
'  wfrm.Caption = "Design Sketch Master Listing"
'  ShowForm wfrm, "RepSM"
'End Sub
'Private Sub RepPM_Click()
'  Dim wfrm As New EmrFrpDsgLst
'  wfrm.Caption = "Design Part Master Listing"
'  ShowForm wfrm, "RepPM"
'End Sub
'Private Sub RepDMCrd_Click()
'  Dim wfrm As New EmrFrpDsgCrd
'  wfrm.Caption = "Design Card Details"
'  ShowForm wfrm, "RepDMCrd"
'End Sub
'Private Sub RepSMCrd_Click()
'  Dim wfrm As New EmrFrpDsgCrd
'  wfrm.Caption = "Design Sketch Card Details"
'  ShowForm wfrm, "RepSMCrd"
'End Sub
'Private Sub RepPMCrd_Click()
'  Dim wfrm As New EmrFrpDsgCrd
'  wfrm.Caption = "Design Part Card Details"
'  ShowForm wfrm, "RepPMCrd"
'End Sub
Private Sub RepRmRt_Click()
  Call FormSelect("RepRmRt")
End Sub
Private Sub RepLabRt_Click()
  Call FormSelect("RepLabRt")
End Sub
Private Sub RepRmRtHist_Click()
  Call FormSelect("RepRmRtHist")
End Sub
Private Sub RepLoc_Click()
  Call FormSelect("RepLoc")
End Sub
Private Sub RepDtTbl_Click()
  Call FormSelect("RepDtTbl")
End Sub
Private Sub RepBnk_Click()
  Call FormSelect("RepBnk")
End Sub
Private Sub RepBagLst_Click()
  Call FormSelect("RepBagLst")
End Sub
Private Sub RepBv_Click()
  Call FormSelect("RepBv")
End Sub
Private Sub RepBagPrn_Click()
  Call FormSelect("RepBagPrn")
End Sub
Private Sub RepDetOrd_Click()
  Call FormSelect("RepDetOrd")
End Sub
Private Sub RepAvgOrd_Click()
  Call FormSelect("RepAvgOrd")
End Sub
Private Sub RepOrdCat_Click()
  Call FormSelect("RepOrdCat")
End Sub
Private Sub RepOrdProfitMrg_Click()
  Call FormSelect("RepOrdProfitMrg")
End Sub
Private Sub RepRmReqStat_Click()
  Call FormSelect("RepRmReqStat")
End Sub
Private Sub RepRmReqDyn_Click()
  Call FormSelect("RepRmReqDyn")
End Sub
Private Sub RepRmReqLoc_Click()
  Call FormSelect("RepRmReqLoc")
End Sub
Private Sub RepBagHist_Click()
  Call FormSelect("RepBagHist")
End Sub
Private Sub RepPrd_Click()
  Call FormSelect("RepPrd")
End Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 25/11/2003  ########################
' Change: Added for Production Points Master Listing
Private Sub RepPrdPts_Click()
  Call FormSelect("RepPrdPts")
End Sub
' ###########################################  Manoj  ###########################################
Private Sub RepWipBPcs_Click()
  Call FormSelect("RepWipBPcs")
End Sub
Private Sub RepWipFgBPcs_Click()
  Call FormSelect("RepWipFgBPcs")
End Sub
Private Sub RepWipBRm_Click()
  Call FormSelect("RepWipBRm")
End Sub

'*** (Jen 2.14)
Private Sub RepWipBRmBkDt_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepWipBRmBkDt")
  End If
  '*** (Jen 2.14 Next)

  'Call FormSelect("RepWipBRmBkDt")   '*** (Bef 2.14 Next)
End Sub
'*** (Jen 2.14)

Private Sub RepWipFgBRm_Click()
  Call FormSelect("RepWipFgBRm")
End Sub
Private Sub RepWipBPcsBkDt_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepWipBPcsBkDt")
  End If
  '*** (Jen 2.14 Next)

  'Call FormSelect("RepWipBPcsBkDt")    '*** (Bef 2.14 Next)
End Sub
Private Sub RepWipFgBPcsBkDt_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepWipFgBPcsBkDt")
  End If
  '*** (Jen 2.14 Next)

  'Call FormSelect("RepWipFgBPcsBkDt")    '*** (Bef 2.14 Next)
End Sub
Private Sub RepRejAna_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepRejAna")
  End If
  '*** (Jen 2.14 Next)
  
  'Call FormSelect("RepRejAna")    '*** (Bef 2.14 Next)
End Sub
' ***** Manali 3.6.0 - 01/09/09 - Gold Loss Report Name and Menu Code Changed To Metal Loss For Users
Private Sub RepMetLs_Click()
  If WithinRTCRange = True Then
    Call FormSelect("RepMetLs")
  End If
End Sub
' ***** Manali 3.6.0 - 01/09/09 - Gold Loss Report Name and Menu Code Changed To Metal Loss For Users
'''Private Sub RepGldLs_Click()
'''  '*** (Jen 2.14 Next)
'''  If WithinRTCRange = True Then
'''    Call FormSelect("RepGldLs")
'''  End If
'''  '*** (Jen 2.14 Next)
'''
'''  'Call FormSelect("RepGldLs")    '*** (Bef 2.14 Next)
'''End Sub

Private Sub RepBrkMsgRm_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepBrkMsgRm")
  End If
  '*** (Jen 2.14 Next)
  
  'Call FormSelect("RepBrkMsgRm")   '*** (Bef 2.14 Next)
End Sub
Private Sub RepBrkMsgPrd_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepBrkMsgPrd")
  End If
  '*** (Jen 2.14 Next)

  'Call FormSelect("RepBrkMsgPrd") '*** (Bef 2.14 Next)
End Sub
Private Sub RepSalAnaDsg_Click()
  Call FormSelect("RepSalAnaDsg")
End Sub
Private Sub RepSalAnaRm_Click()
  Call FormSelect("RepSalAnaRm")
End Sub
Private Sub RepSalAnaLab_Click()
  Call FormSelect("RepSalAnaLab")
End Sub
Private Sub RepDiaProfit_Click()
  Call FormSelect("RepDiaProfit")
End Sub
Private Sub RepSOPnd_Click()
  Call FormSelect("RepSOPnd")
End Sub
Private Sub RepWIWPrd_Click()
  Call FormSelect("RepWIWPrd")
End Sub
Private Sub RepWIWExp_Click()
  Call FormSelect("RepWIWExp")
End Sub
Private Sub RepWIWLongPrd_Click()
  Call FormSelect("RepWIWLongPrd")
End Sub
Private Sub RepWIWLongExp_Click()
  Call FormSelect("RepWIWLongExp")
End Sub
Private Sub RepStkLed_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepStkLed")
  End If
  '*** (Jen 2.14 Next)

  'Call FormSelect("RepStkLed")   '*** (Bef 2.14 Next)
End Sub
Private Sub RepFlashStk_Click()
  Call FormSelect("RepFlashStk")
End Sub
Private Sub RepCustmStk_Click()
  Call FormSelect("RepCustmStk")
End Sub
Private Sub RepStkSumm_Click()
  '*** (Jen 2.14 Next)
  If WithinRTCRange = True Then
    Call FormSelect("RepStkSumm")
  End If
  '*** (Jen 2.14 Next)
  
  'Call FormSelect("RepStkSumm")   '*** (Bef 2.14 Next)
End Sub
Private Sub RepTxnDT_Click()
  Call FormSelect("RepTxnDT")
End Sub
Private Sub RepTxnBS_Click()
  Call FormSelect("RepTxnBS")
End Sub
Private Sub RepTxnMLT_Click()
  Call FormSelect("RepTxnMLT")
End Sub
Private Sub RepTxnBV_Click()
  Call FormSelect("RepTxnBV")
End Sub
Private Sub RepTxnREJ_Click()
  Call FormSelect("RepTxnREJ")
End Sub
Private Sub RepTxnPR_Click()
  Call FormSelect("RepTxnPR")
End Sub
Private Sub RepTxnMV_Click()
  Call FormSelect("RepTxnMV")
End Sub
Private Sub RepTxnCNV_Click()
  Call FormSelect("RepTxnCNV")
End Sub
Private Sub RepUnMtchCNV_Click()
  Call FormSelect("RepUnMtchCnv")
End Sub
'Private Sub RepTxnREC_Click()
'  Dim wfrm As New EmrFrpTxnLst
'  wfrm.Caption = "Recovery Transaction Listing"
'  ShowForm wfrm, "RepTxnREC"
'End Sub
Private Sub RepFbLst_Click()
  Call FormSelect("RepFbLst")
End Sub
Private Sub RepFwbLst_Click()
  Call FormSelect("RepFwbLst")
End Sub
Private Sub RepFgMvm_Click()
  Call FormSelect("RepFgMvm")
End Sub

'Private Sub RepFrLst_Click()
'  Dim wfrm As New EmrFrpFgList
'  wfrm.Caption = "Finished Goods (RM)"
'  ShowForm wfrm, "RepFrLst"
'End Sub
'Private Sub RepFwrLst_Click()
'  Dim wfrm As New EmrFrpFgList
'  wfrm.Caption = "Finished Goods (RM Return)"
'  ShowForm wfrm, "RepFwrLst"
'End Sub
Private Sub RepVchPrnDT_Click()
  Call FormSelect("RepVchPrnDT")
End Sub
Private Sub RepVchPrnBS_Click()
  Call FormSelect("RepVchPrnBS")
End Sub
Private Sub RepVchPrnMLT_Click()
  Call FormSelect("RepVchPrnMLT")
End Sub
Private Sub RepVchPrnBV_Click()
  Call FormSelect("RepVchPrnBV")
End Sub
Private Sub RepVchPrnREJ_Click()
  Call FormSelect("RepVchPrnREJ")
End Sub
Private Sub RepVchPrnPR_Click()
  Call FormSelect("RepVchPrnPR")
End Sub
Private Sub RepVchPrnMV_Click()
  Call FormSelect("RepVchPrnMV")
End Sub
Private Sub RepVchPrnCNV_Click()
  Call FormSelect("RepVchPrnCNV")
End Sub
'Private Sub RepVchPrnREC_Click()
'  Dim wfrm As New EmrFrpVchPrn
'  wfrm.Caption = "Recovery Voucher Printing"
'  ShowForm wfrm, "RepVchPrnREC"
'End Sub
'Private Sub RepMemIss_Click()
'  Dim wfrm As New EmrFrpMemoIss
'  wfrm.Caption = "Memo Issue"
'  ShowForm wfrm, "RepMemIss"
'End Sub
Private Sub RepDetInv_Click()
  Call FormSelect("RepDetInv")
End Sub
Private Sub RepAvgInv_Click()
  Call FormSelect("RepAvgInv")
End Sub
Private Sub RepInvCat_Click()
  Call FormSelect("RepInvCat")
End Sub
Private Sub RepInv_Click()
  Call FormSelect("RepInv")
End Sub
Private Sub RepGrFormFrnt_Click()
  Call FormSelect("RepGrFormFrnt")
End Sub
Private Sub RepGrFormBack_Click()
  Call FormSelect("RepGrFormBack")
End Sub

'*** (Jen 2.14)
Private Sub RepGRFormBackMulti_Click()
  Call FormSelect("RepGRFormBackMulti")
End Sub
'*** (Jen 2.14)

Private Sub RepShpBill_Click()
  Call FormSelect("RepShpBill")
End Sub
'*** (Jen 2.13)
Private Sub RepGrFormFrntMulti_Click()
  Call FormSelect("RepGrFormFrntMulti")
End Sub
Private Sub RepShpBillMulti_Click()
  Call FormSelect("RepShpBillMulti")
End Sub
'*** (Jen 2.13)
Private Sub RepPckLst_Click()
  Call FormSelect("RepPckLst")
End Sub
Private Sub RepPckLstInv_Click()
  Call FormSelect("RepPckLstInv")
End Sub
Private Sub RepInvBOELst_Click()
  Call FormSelect("RepInvBOELst")
End Sub
Private Sub RepGSPFORM_Click()
  Call FormSelect("RepGSPFORM")
End Sub
Private Sub RepGSPDetLet_Click()
  Call FormSelect("RepGSPDetLet")
End Sub
Private Sub RepGSPCert_Click()
  Call FormSelect("RepGSPCert")
End Sub
Private Sub RepDeclnLet_Click()
  Call FormSelect("RepDeclnLet")
End Sub

'*** (Jen 3.01)
Private Sub RepGSPFORMMulti_Click()
  Call FormSelect("RepGSPFORMMulti")
End Sub
Private Sub RepGSPDetLetMulti_Click()
  Call FormSelect("RepGSPDetLetMulti")
End Sub
Private Sub RepGSPCertMulti_Click()
  Call FormSelect("RepGSPCertMulti")
End Sub
Private Sub RepDeclnLetMulti_Click()
  Call FormSelect("RepDeclnLetMulti")
End Sub
'*** (Jen 3.01)

Private Sub RepTagPrn_Click()
  Call FormSelect("RepTagPrn")
End Sub
Private Sub RepInvProfitMrg_Click()
  Call FormSelect("RepInvProfitMrg")
End Sub
Private Sub RepInvLossSubRep_Click()
  Call FormSelect("RepInvLossSubRep")
End Sub
Private Sub RepAnnexValAddn_Click()
  Call FormSelect("RepAnnexValAddn")
End Sub
Private Sub RepValAddn_Click()
  Call FormSelect("RepValAddn")
End Sub
Private Sub PpcCap_Click()
  Call FormSelect("PpcCap")
End Sub
Private Sub PpcDsgLd_Click()
  Call FormSelect("PpcDsgLd")
End Sub
Private Sub PpcOrdPln_Click()
  Call FormSelect("PpcOrdPln")
End Sub

'****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master] ******
Private Sub VldRmLn_Click()
  Call FormSelect("VldRmLn")
End Sub
'****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master] ******

'****** Manali 3.3.0 (04/07/08) - RmSz Desc Entry
Private Sub RmIdSz_Click()
  Call FormSelect("RmIdSz")
End Sub
'****** Manali 3.3.0 (04/07/08) - RmSz Desc Entry


' Zubin 213
'Private Sub XMLMstToEmr_Click()
'  Call FormSelect("XMLMstToEmr")
'End Sub
' Zubin 213
'Private Sub XMLToEmr_Click()
'  Call FormSelect("XMLToEmr")
'End Sub
Private Sub Zoom_Click()
  Call FormSelect("Zoom")
End Sub
Private Sub LsRec_Click()
  Call FormSelect("LsRec")
End Sub
Private Sub RepLsRec_Click()
  Call FormSelect("RepLsRec")
End Sub
Private Sub RepCustRm_Click()
  Call FormSelect("RepCustRm")
End Sub
'*** Added Again In 3.2.0
Private Sub RepCustmStkLand_Click()
  '*** Jay 3.1.1 (All Rates Rpt)
  Call FormSelect("RepCustmStkLand")
End Sub
'*** Added Again In 3.2.0
Private Sub RepGrFormFrntNew_Click()
 Call FormSelect("RepGrFormFrntNew")
End Sub
'*** Added Again In 3.2.0
Private Sub RepGrFormFrntNewMul_Click()
'*** Jay 3.1.1
  Call FormSelect("RepGrFormFrntNewMul")
End Sub

'*** (Jen 2.12)
Private Sub RepInvLoose_Click()
  Call FormSelect("RepInvLoose")
End Sub
Private Sub RepShpBillLoose_Click()
  Call FormSelect("RepShpBillLoose")
End Sub
Private Sub RepGrFormLoose_Click()
  Call FormSelect("RepGrFormLoose")
End Sub
Private Sub RepPckLstLoose_Click()
  Call FormSelect("RepPckLstLoose")
End Sub
'*** (Jen 2.12)
'U.jEmr.17- new menu added to Voucher profit margin report
Private Sub JPRFMRGN_Click()
  Call FormSelect("JPRFMRGN")
End Sub

'PG.7 new menu added For Design Flash Stock Report
Private Sub RepJtDsgFlashStk_Click()
Call FormSelect("RepJtDsgFlashStk")
End Sub

Private Sub RepJtTxnLstRpt_Click()
Call FormSelect("RepJtTxnLstRpt")
End Sub

' ########################  Manoj #### Ver: 2.0.6 #### Date: 17/02/2004  ########################
Private Sub UsrMail_Click()
  Call FormSelect("UsrMail")
End Sub
' ###########################################  Manoj  ###########################################

' ########################  Manoj #### Ver: 2.11.0 #### Date: 27/05/2005  ########################
Private Sub MailSignIn_Click()
  Call MailLogin
End Sub
Private Sub SNDMAIL_Click()
  If gb_MailAct = True Then
    If ComposeMsg = True Then DispMsg "Mail Successfully Sent to OutBox.", etInfo
  Else
    DispMsg "Mail Login is Not Available.", etWarning
  End If
End Sub
Private Sub MailSignOut_Click()
  Call LogOffSes
  DispMsg "Mail Sign Out Succeeded.", etInfo
End Sub
' ########################  Manoj #### Ver: 2.11.0 #### Date: 27/05/2005  ########################

Private Function GetVersion(ByRef ReturnMsg As String) As Boolean
  Dim wSqlStrg  As String
  GetVersion = False
  Dim wRsVer As MwfLib.MDORowSet, wExeVer As String, wDatVer As String
  
  wExeVer = App.Major
  wExeVer = wExeVer & "." & App.Minor
  wExeVer = wExeVer & "." & App.Revision

  wSqlStrg = "Select PDesc225 From Param where PTyp= 'VER' AND PMCD ='DATVER'"
  Set wRsVer = ADC.Connection.OpenRes(wSqlStrg)
  If wRsVer.RecCount <> 0 Then
    wDatVer = wRsVer.FldValue("PDesc225")
  Else
    ReturnMsg = "Cannot proceed. No Version Record found in Param"
    GetVersion = False
    Exit Function
  End If
  
  If wDatVer = wExeVer Then
    GetVersion = True
    ReturnMsg = wDatVer
  Else
    ReturnMsg = "Cannot proceed. Exe Version is " & wExeVer & Chr(13) & "Database Version is " & wDatVer
    GetVersion = False
  End If
End Function

Private Function OldGetVersion() As String
  Dim wSqlStrg  As String
  '*** Jen ***
  '*** Use this code to compare the Exe Version with the Database Version
  '*** Stop the user from accessing the System if there is a version mismatch between the
      'Exe and the Database

  'GetVersion = True
  OldGetVersion = ""
  Dim wRsVer As MwfLib.MDORowSet, wExeVer As String, wDatVer As String
  'If UCase(App.EXEName) = "EMR"
  wExeVer = App.Major
  wExeVer = wExeVer & "." & App.Minor
  wExeVer = wExeVer & "." & App.Revision

  wSqlStrg = "Select PDesc225 From Param where PTyp= 'VER' AND PMCD ='DATVER'"
  Set wRsVer = ADC.Connection.OpenRes(wSqlStrg)
  If wRsVer.RecCount <> 0 Then wDatVer = wRsVer.FldValue("PDesc225")
  If wDatVer = wExeVer Then
    'GetVersion = True
    OldGetVersion = wDatVer
  Else
    'GetVersion = False
    MsgBox "Exe Version is " & wExeVer & Chr(13) & "Database Version is " & wDatVer
  End If
  '*** Jen ***
End Function

'****** Sachin 3.3.0 (01-07-08) - [16. Check Modules and Ini Users]******
Private Function CheckIntegrity(ByVal pi_IniUsers As Integer) As String

Dim wo_rsChkInt As MwfLib.MDORowSet, wiCtr As Integer, ws_ModulesFldStr As String

'Ini Users Check
CheckIntegrity = ""

If Not ADC.Connection.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + Trim(TxtCoCd.text) + "' ") Then _
  CheckIntegrity = "Invalid Company Code/Company Code Does Not Exist In Head Master": Exit Function

' ***** Manali 3.7.1 - JMIC added in following List
If pi_IniUsers < ADC.Connection.GetFldVal("Select Count('x') From Param Where PTyp='USR' And PValidYN='Y' And PMCd Not In ('MIC','SUPP','CRM','ADM', 'JMIC')") Then _
  CheckIntegrity = "Database Users are more than the Users Specified In the Ini File": Exit Function
  
Call SetModuleSeq   'Add every new module to the array in this Subroutine in EmrMod
  
ws_ModulesFldStr = ""
Do While wiCtr <= UBound(gs_ModulesStr)
  ws_ModulesFldStr = ws_ModulesFldStr + "+" + gs_ModulesStr(wiCtr)
  wiCtr = wiCtr + 1
Loop

'Modules Check
Set wo_rsChkInt = ADC.Connection.OpenResultset(" Select hChkIntgty, (HCoCd+':'+HName+':'" + ws_ModulesFldStr + ") as qModStatus " + _
                                               "  From Head " + _
                                               " Where HCd='" + ctSelfCmCd + "' And HCoCd='" + Trim(TxtCoCd.text) + "'")
With wo_rsChkInt
  If Not (.BOF Or .EOF) Then _
        If Decrypt(!hChkIntgty, GenCryptNew) <> !qModStatus Then _
        CheckIntegrity = "Modules Check Failed in the Current Database, Contact Microway!!"
End With

End Function
'****** Sachin 3.3.0 (01-07-08) - [16. Check Modules and Ini Users]******

Public Function DtAndCoChk() As Boolean
'*** Jen ***
'****** The max date till which the exe can be used and the company name are stored in the
       'emr.ini file in an encrypted form. It is the third parameter in the Emr.ini file.
       '(list of parameters are separated by comma)
  Dim wSqlStrg  As String
  DtAndCoChk = False
  
  On Error GoTo ErrH
  Dim wFlNm As String, wDtAndCoChk As Variant, wDt As String
  Dim wCompanyNm As String, wHName As String, wRsHd As MwfLib.MDORowSet
  
  '****** Sachin 2.12 [Date Check through SQL]
  Dim wb_DateChk As Boolean, wb_YyChk As Boolean
  '****** Sachin 2.12 [Date Check through SQL]
  
  '****** If the first character in the max date is 0, it is not considered while decrypting.
  'So '01/10/02' is decrypted as '01/10/02'
'  wDtAndCoChk = MWLib.DtAndCoChkNo
  If Len(wDtAndCoChk) >= 8 Then
    wDt = IIF(InStr(1, wDtAndCoChk, "/") = 2, Mid(wDtAndCoChk, 1, 7), Mid(wDtAndCoChk, 1, 8))
    wCompanyNm = IIF(InStr(1, wDtAndCoChk, "/") = 2, Mid(wDtAndCoChk, 8), IIF(Len(wDtAndCoChk) > 8, Mid(wDtAndCoChk, 9), ""))
  End If

'****** If the Max Date or Company Name is missing in the Emr.ini file then an error message is raised
  If Not IsDate(wDt) Then Err.Raise vbObjectError + 512, "", "Date Dets Missing/Improper In Ini File"
  If wCompanyNm = "" Then Err.Raise vbObjectError + 512, "", "Company Dets Missing in Ini file"

'****** Get Company Name form Head file and check if it is the same as that in the Ini file
'****** Check if the System Date is less than or equal to the date in the Ini file
'****** If any of the above conditions are false then the Datecheck has failed
  wSqlStrg = "Select HName From Head where HCoCd= '" + TxtCoCd + "' and HCd='" + ctSelfCmCd + "'"
  Set wRsHd = ADC.Connection.OpenRes(wSqlStrg)
  If wRsHd.RecCount <> 0 Then wHName = wRsHd.FldValue("HName")
  ' Before 15/02/06
  'If ADC.Connection.SrvrDate > CDate(wDt) Then Err.Raise vbObjectError + 512, "", "Get Updated Exe. Contact Microway"
  
  'Zubin 212 (Bef 05/09/06)
  'If ADC.Connection.SrvrDate > CDate(wDt) Then Err.Raise vbObjectError + 512, "", "Master Db Sql Error: 251001. Contact System Administrator."
  
  ' Zubin 212 (05/09/06)
  Dim wi_IniUBnd As Integer
  wi_IniUBnd = MWLib.NoOfUsers * 10
  
  '****** Sachin - Temporary Change-05/03/2007  [300 Days extension to the Ini Date]******
  wDt = CDate(wDt) '+ 300 - Extension Days Commented in Emr 3.02.0
  If ADC.Connection.GetFldVal("Set DateFormat DMY Select Case When Cast((Convert(VarChar(10),getdate(),103)) As SmallDateTime)<=Cast('" + wDt + "' As SmallDateTime) Then 'Y' Else 'N' End") = "Y" And _
        ADC.Connection.GetFldVal("Select PValue3 From Param Where PTyp= 'USR' and PMCd= 'MIC'") <> "   " Then
      If Not ADC.Connection.RecSeek("Select 'x' From Param Where PTyp='USR' And PMCd='MIC' And ModTime=0") Then _
        ADC.Connection.Execute ("Set DateFormat DMY Update Param set PNum1= 99, PValue3=Space(3), ModTime= 0 Where PTyp= 'USR' and PMCd= 'MIC'")
  End If
  '****** Sachin - Temporary Change-05/03/2007 [60 Days extension to the Ini Date]******
  
  '****** [Sachin 2.12 - 05-10-2006] Change of Date Check
  '****** Additional check [wb_yychk] to check if the ini year is greater than the current year of server + 3
  wb_DateChk = IIF(ADC.Connection.GetFldVal("Set DateFormat DMY Select Case When Cast((Convert(VarChar(10),getdate(),103)) As SmallDateTime)>Cast('" + wDt + "' As SmallDateTime) Then 'Y' Else 'N' End") = "Y", True, False)
  wb_YyChk = IIF(ADC.Connection.GetFldVal("Set DateFormat DMY Select Case When Year(getdate())+3<Year(Cast('" + wDt + "' As SmallDateTime)) Then 'Y' Else 'N' End") = "Y", True, False)
  '****** [Sachin 2.12 - 05-10-2006] Change of Date Check
 
 '**** [Sachin 2.12 - 05-10-2006] Commented Below is the Earlier Check - It Gave 'Contact Microway' error if any client year range was less than the Ini year
 ' If ADC.Connection.SrvrDate > CDate(wDt) Or ADC.Connection.GetFldVal("Select ModTime from Param where PTyp= 'USR' and PMCd= 'MIC' ") <> 0 Then
 '**** [Sachin 2.12 - 05-10-2006] Commented Below is the Earlier Check - It Gave 'Contact Microway' error if any client year range was less than the Ini year

  If wb_DateChk = True Or ADC.Connection.GetFldVal("Select ModTime from Param where PTyp= 'USR' and PMCd= 'MIC' ") <> 0 Then
    If wb_DateChk = True Then ADC.Connection.Execute ("Update Param set PNum1= 99, ModTime= ModTime + (Case When ModTime>= 999 Then 0 Else 1 End) where PTyp= 'USR' and PMCd= 'MIC'")
    Randomize
    If ADC.Connection.GetFldVal("Select ModTime from Param where PTyp= 'USR' and PMCd= 'MIC' ") > Int(((wi_IniUBnd - 10) * Rnd) + 11) Or _
        ADC.Connection.GetFldVal("Select ModTime from Param where PTyp= 'USR' and PMCd= 'MIC' ") >= 999 Then
      RestoreColors
      Call SetRecClrSet(Nothing)
      Call SetMDIFrm(Nothing)
      Set mo_MDIPic = Nothing: Set mo_CsTlBarPic = Nothing: Set mo_TlBarColl = Nothing
      End
    End If
  End If
  ' Zubin 212 (05/09/06)
  
  '****** [Sachin 22-1-2007] Ini Check updated especially for Elegant Coll.
  If wb_YyChk = True Then Err.Raise vbObjectError + 512, "", "Improper Ini File, Please Update"
  '****** [Sachin 22-1-2007] Ini Check updated especially for Elegant Coll.
  
  '*** Jay 2.14 ***
  If UCase(wCompanyNm) <> UCase(wHName) Then Err.Raise vbObjectError + 512, "", "Company Name Not Matching"
  'If wCompanyNm <> wHName Then Err.Raise vbObjectError + 512, "", "Company Name Not Matching"
  '*** Jay 2.14 ***

'****** Done

  DtAndCoChk = True
  Exit Function

ErrH:
  DtAndCoChk = False
  'Zubin 212 (05/09/06) [Err.Description chk added]
  If Err.Description <> "" Then MsgBox Err.Description
'*** Jen ***
End Function
Public Function FnCoCdChk() As Boolean
'*** Jen ***
  FnCoCdChk = False
  On Error GoTo ErrH
  Dim ws_CoCd As String

  If Trim(TxtCoCd) = "" Then Err.Raise vbObjectError + 512, "", "Company Code Should Be Entered"
  ws_CoCd = ADC.Connection.GetFldVal("Select HCoCd From Head where HCoCd= '" + TxtCoCd + "'")
  If UCase(ws_CoCd) <> UCase(TxtCoCd) Then Err.Raise vbObjectError + 512, "", "Company Code Not Matching"

  FnCoCdChk = True
  Exit Function

ErrH:
  FnCoCdChk = False
  MsgBox Err.Description
'*** Jen ***
End Function
Public Function ValidNetId() As Boolean
'  '*** Jen ***
'  '*** This is the function which has to be called to check if the NetWork Card Number
'      'of the database server is valid.
'  '*** The number which is passed as the parameter of DecryptPwd is the number
'      'which is obtained after encrypting the Lan Card Number of the Database Server.
'      'It can be generated by entering the NetWork Card number of the Database Server
'      'in the PrjGetNo exe
'      'This is stored in the 'Emr.ini' file.
'      '(separated by a comma from the other values in the ini file)
'      'This number is different for different companies (So it has to
'      'be changed in the Emr.ini file for each company)
'  '*** If the encoded Network Card number of the Database Server is different
'      'than the one in the Emr.ini file then display an error message 'invalid install'
'
'
'  Dim wNetId As String
'  Dim wIniId As String
'  Dim NetIdChk As Object
'  Dim wlen As Integer
'  Set NetIdChk = CreateObject("mwrlib.netidchk")
'  wNetId = NetIdChk.IdNo
'  Set NetIdChk = Nothing
'
'  'wIniId = DecryptPwd("72212.651650900.477")
'  wIniId = MWLib.NetWrkCrdNo
'  wlen = Len(wIniId)
'  If Right(wIniId, wlen) = Right(wNetId, wlen) Then
'      ValidNetId = True
'  Else
'      ValidNetId = False
'  End If
'  '*** Jen ***
End Function

Private Sub CmdOk_Click()
    MWLib.BeginProcess Me, "Loading " + Me.Caption
    Me.MousePointer = vbHourglass
    Dim wo_rsUsr As MwfLib.MDORowSet, wo_cn As New MwfLib.MDOConnection
    Dim wSqlStrg As String, ws_Tmp As String, wChkIntegrity As String
    Dim wMultiComp As Boolean, wSelfCocdNm As String
 
    Dim wDbVer As String
    
    If Not ChkInstance Then GoTo MDIEnd
    
    If Not MWLib.SetDateFormat Then GoTo MDIEnd
    If Not Ctl.SetLogInDets(TXT_UCD, ctProjectName, TxtCoCd, cmbSrvrNm.text, cmbDbNm.text) Then GoTo MDIEnd
    Set Ctl.MDIApp = Me
    If Not ADC.InitRtn Then GoTo MDIEnd
    On Error GoTo ErrH
    Call Ctl.SetHlpConnection
    If Ctl.GetHlpConnection Is Nothing Then GoTo MDIEnd
    Set go_HlpCn = Ctl.GetHlpConnection
                
    gs_Partition = ctCurrPrtn
    gbSvrChk = IIF(UCase(MWLib.DbSrvrNm) = UCase(MWLib.RepDbSrvrNm), True, False)
    
    '****** Sachin 3.02 24-12-07 - Setting the Purge Date
    'sv.9 getting purge date from vparam
    gs_PurgeDt = GetPurgeDt(TxtCoCd.text, "JT")
                  
    Dim wErrMsg As String
    If CoChk(wErrMsg) = False Then
      MsgBox wErrMsg
      GoTo MDIEnd
    End If
    
    Dim wRetMsg As String, wSucc As Boolean
    wSucc = GetVersion(wRetMsg)
    gs_Ver = ""
    If wSucc Then
      gs_Ver = wRetMsg
    Else
      MsgBox wRetMsg
      GoTo MDIEnd
    End If
    
    If MWLib.IsItLaptop = False And ((MWLib.NoOfUsers) < ADC.Connection.GetFldVal("Select Count('x') From Param Where PTyp='USR' And PValidYN='Y' And PMCd<>'MIC'")) Then
        MsgBox "Database Users are more than the Users Specified In the Ini File"
        GoTo MDIEnd
    End If


    ' Old Way
  '   If Not ResolutionChk Then
  '     MsgBox "Screen Resolution Should Be 1024x768 Or Higher", vbOKOnly, "Screen Resolution"
  '    GoTo MDIEnd
  '   End If
  '  Old
'    If ADC.Connection.RecSeek("Select 'x' From Param where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= '' and PValidYN= 'N'") = True Then
 '     MsgBox "Cannot Login As Database Is In Single User Mode For Some Exclusive Updates"
  '    GoTo MDIEnd
'    End If
    '*** (Jen 3.01)
    Call HideAllMnus
    wSqlStrg = "SELECT * FROM PARAM WHERE PTYP= 'USR' AND PMCD ='" + TXT_UCD.text + "' And PValidYn = 'Y'"

    Dim wl_UsrCnt As Long
    
 '   If wl_UsrCnt >= MWLib.NoOfUsers Then
  '    MsgBox "No of Users Allowed (" & MWLib.NoOfUsers & ") are Already Logged On", vbOKOnly, "User Count Check"
   '   GoTo MDIEnd
    'Else
    
    Set wo_rsUsr = ADC.Connection.OpenRes(wSqlStrg)
    If wo_rsUsr.RecCount = 0 Then
            MsgBox ("Invalid User code")
            TXT_UCD.SetFocus
            GoTo MDIEnd
    End If
    If wo_rsUsr.FldValue("PValue") <> TXT_PWD.text Then
            MsgBox ("Invalid Password")
            TXT_PWD.SetFocus
            GoTo MDIEnd
    End If
     
    Dim wo_rsUL As MwfLib.MDORowSet
            
    gs_HostNm = UCase$(GetHostNm)
    Dim ws_UsrLoggedInHostComputer As String
    Dim wi_delusr As Integer
    '6.1-49 UlSysCd='JEMR' added
    ws_UsrLoggedInHostComputer = ADC.Connection.GetFldVal("Select UlHostNm From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlHostNm <> '" + gs_HostNm + "' and UlSysCd='JEMR'")
    '6.1-49 UlSysCd='JEMR' added
    If ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlHostNm='" + gs_HostNm + "' and UlSysCd='JEMR'") Then
       wi_delusr = MsgBox("User '" & UCase$(TXT_UCD.text) & "' Already Logged In On This Machine. Delete User ?", vbYesNo, "User Check")
       If wi_delusr = vbYes Then
       '6.1-49 UlSysCd='JEMR' added
         ADC.Connection.Execute ("Delete From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' /* and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' */ And UlHostNm='" + gs_HostNm + "' and UlSysCd='JEMR'")
         'GoTo MDIAddUsrRec
       Else
         GoTo MDIEnd
       End If
    '6.1-49 UlSysCd='JEMR' added
    ElseIf ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlHostNm <> '" + gs_HostNm + "' and UlSysCd='JEMR'") Then
          MsgBox "Cannot Login. User '" & UCase$(TXT_UCD.text) & "' Already Logged In On Computer '" + ws_UsrLoggedInHostComputer + "', Company " + UCase$(Trim$(TxtCoCd.text)), vbOKOnly, "User Check"
          GoTo MDIEnd
    End If
        
    wl_UsrCnt = CLng(ADC.Connection.GetFldVal("Select Count(Distinct(UlUsrCd)) from UsrLogin"))
    If wl_UsrCnt >= MWLib.NoOfUsers Then
      MsgBox "No of Users Allowed (" & MWLib.NoOfUsers & ") are Already Logged On", vbOKOnly, "User Count Check"
      GoTo MDIEnd
    End If
    
MDIAddUsrRec:   '6.1-49 UlSysCd='JEMR' added
        Set wo_rsUL = ADC.Connection.OpenResultset("Select * from UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlSysCd='JEMR'")
        If Not (wo_rsUL Is Nothing) Then
          With wo_rsUL
            .Addnew
            .FldValue("UlUsrCd") = UCase$(Trim$(TXT_UCD.text))
            .FldValue("UlCoCd") = UCase$(Trim$(TxtCoCd.text))
            .FldValue("UlHostNm") = gs_HostNm
            .FldValue("UlDt") = ADC.Connection.SrvrDate
            .FldValue("UlTime") = ADC.Connection.SrvrTime
            .FldValue("UlSysCd") = "JEMR"    '6.1-49
            .Update
          End With
      End If
'    End If
    '6.1-49 UlSysCd='JEMR' added
    gs_UlIdNo = ADC.Connection.GetFldVal("Select UlIdNo From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' " + _
                                        " and UlHostNm='" + gs_HostNm + "' and UlSysCd='JEMR'")
      Dim ws_ClrCd As String
      Dim wo_Pic1 As Picture, wo_Pic2 As Picture, wo_PicAdc As Picture, wb_PicExists As Boolean
      
      '*** (Jen 2.13)
      wb_ApplyClr = False
      
      ws_ClrCd = ADC.Connection.GetFldVal("Select PValue2 from Param where PTyp= 'USR' and PMCd= '" + TXT_UCD + "'")
      Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
      
      If Not (GetRecClrSet.RecCount > 0) Then
        ws_ClrCd = ADC.Connection.GetFldVal("Select PValue From Param where PTyp= 'DEFCLR'")
        If ws_ClrCd <> "" Then Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
      End If
      
      If GetRecClrSet.RecCount > 0 Then wb_ApplyClr = True
      
      If wb_ApplyClr = True Then
      '*** (Jen 2.13)
      
        ws_ClrCd = ADC.Connection.GetFldVal("Select PValue2 from Param where PTyp= 'USR' and PMCd= '" + TXT_UCD + "'")
        Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
                      
        
              
              If Not (GetRecClrSet.RecCount > 0) Then
                ws_ClrCd = ADC.Connection.GetFldVal("Select PValue From Param where PTyp= 'DEFCLR'")
                Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
              End If
              
              Call SetClrPth("")
              If InStrRev(App.Path, "\") > 1 Then Call SetClrPth(Mid(App.Path, 1, InStrRev(App.Path, "\") - 1))
              If GetClrPth <> "" Then Call SetClrPth(GetClrPth & "\MwClrSet\")
              Picture1.Visible = False   '*** 15/06/05
              
              wb_PicExists = True
              
              If Trim(GetRecClrSet!CsADCBtnPic) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsADCBtnPic) = False Or Trim(GetRecClrSet!CsBtnPicSet1) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsBtnPicSet1) = False Or Trim(GetRecClrSet!CsBtnPicSet2) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsBtnPicSet2) = False Or Trim(GetRecClrSet!CsFormPic) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsFormPic) = False Then
                wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_PicAdc = LoadPicture("")    '*** This will assign 0 to woPic
                Set wo_PicAdc = LoadPicture(GetClrPth & GetRecClrSet!CsADCBtnPic)
                If (wo_PicAdc Is Nothing) Or wo_PicAdc = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_Pic1 = LoadPicture("")
                Set wo_Pic1 = LoadPicture(GetClrPth & GetRecClrSet!CsBtnPicSet1)
                If (wo_Pic1 Is Nothing) Or wo_Pic1 = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_Pic2 = LoadPicture("")
                Set wo_Pic2 = LoadPicture(GetClrPth & GetRecClrSet!CsBtnPicSet2)
                If (wo_Pic2 Is Nothing) Or wo_Pic2 = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set Picture1.Picture = LoadPicture("")
                Set Picture1.Picture = LoadPicture(GetClrPth & GetRecClrSet!CsFormPic)
                If (Picture1.Picture Is Nothing) Or Picture1.Picture = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = False Then
                ws_ClrCd = ctDefClrCd
                Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
              End If
            
            '*** (Jen 2.13)
            Else
              RestoreColors
              Picture1.Visible = False
              Me.Picture = LoadPicture("")
              Call SetRecClrSet(Nothing)
              Set mo_MDIPic = Nothing
              Set mo_CsTlBarPic = Nothing
            End If
            '*** (Jen 2.13)
            
            SetMDISysColorsAfterLogin
            '*** (Jen 2.13)
            If Not (GetRecClrSet Is Nothing) Then
            '*** (Jen 2.13)
              If GetRecClrSet.RecCount > 0 Then
                GetRecClrSet.MoveFirst
                If Trim(GetRecClrSet!CsMDIPic) <> "" And IsValidPath(GetClrPth & GetRecClrSet!CsMDIPic) = True Then
                  Set mo_MDIPic = LoadPicture(GetClrPth & GetRecClrSet!CsMDIPic)
                End If
                If Not (mo_MDIPic Is Nothing) Then Set PictureClip1.Picture = mo_MDIPic
                SetMDIPic
                Me.BackColor = Val(GetRecClrSet!CsLblForClr)
              End If
            
            '*** (Jen 2.13)
            End If
            '*** (Jen 2.13)
            '*** Jenny Color

            gs_UsrCd = TXT_UCD
            WindowState = 2

            ' Zubin 212 (27/02/06)
            Dim wo_RsModules As MwfLib.MDORowSet
            Dim ws_Mod As String
            ' **** Manali 3.10.0 - 04/04/12 - Multi Price Quotation Module
            Set wo_RsModules = ADC.Connection.OpenRes("select HCoCd, HPPCYN, HCtbMod, HFgQuotYN, HLooseInv, " + _
                               "HConsolidatedInv, hAutomateReturns, /* HLclInvYn, HFgTagMod,*/ HCustomiseOrd, HTravelExlYN, HDsgCatYN, HRfIdMod, HMultiPrcQtMod from Head " + _
                               "where HCoCd ='" + UCase(TxtCoCd.text) + "' and HCd= '" + ctSelfCmCd + "'")
            If UCase(wo_RsModules.FldValue("HPPCYN")) = "N" Then ws_Mod = ws_Mod + "'PPC','PPCCAP','PPCDSGLD','PPCORDPLN','REPPPC','REPDSGLD','REPCNTRLCPY','REPORDPLN'"
            If UCase(wo_RsModules.FldValue("HFgQuotYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'QB'"
            If UCase(wo_RsModules.FldValue("HLooseInv")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepInLoose'"   '"'RepInLoose', 'RepInvLoose', 'RepShpBillLoose', 'RepGrFormLoose', 'RepPckLstLoose'"
            If UCase(wo_RsModules.FldValue("HConsolidatedInv")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepGrFormFrntMulti', 'RepShpBillMulti', 'RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti', 'RepGrFormFrntNewMul', 'RepGRFormBackMulti'"
            If UCase(wo_RsModules.FldValue("HAutomateReturns")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RETMEM'"
            If UCase(wo_RsModules.FldValue("HCustomiseOrd")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'IdeDfn', 'IdeUsrRight'"
            If UCase(wo_RsModules.FldValue("HTravelExlYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepTravExcel'"
            If UCase(wo_RsModules.FldValue("HDsgCatYn")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'DsgPrm'/*, 'DsgPrmCat'*/ "
            If UCase(wo_RsModules.FldValue("HRfIdMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RfIdMst'"
            If UCase(wo_RsModules.FldValue("HMultiPrcQtMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'MultiPrcQt', 'RepMultiPrcQt'"
            If Not ADC.Connection.RecSeek("select 'x' from SysDb where SysId= 'CRM' ") Then
              ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'SALORDACK'"
            End If
            ws_Mod = IIF(ws_Mod <> "", " and UaMnuCd Not In (" + ws_Mod + ")", "")
            Set wo_RsModules = Nothing
            
            Dim wo_rsusracc As MwfLib.MDORowSet

            ' Manoj 2.10.0
            Dim wsCoCd As String
            ' Zubin 212 (27/02/06) ( + ws_Mod ADDED)
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from UsrAccess where uausrcd='" + UCase$(TXT_UCD.text) + "' And UaMnuForm = '" + ms_MnuForm + "'" + _
                      " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.text) & ",%' End" + ws_Mod)
            ' Manoj 2.10.0
            
            '*** Jenny Color (10/08/05)
            Call CreateToolBar
            Dim wi_i As Integer
            wi_i = 0
            '*** Jenny Color (10/08/05)
            
            Do While Not (wo_rsusracc.EOF Or wo_rsusracc.BOF)
              '****** Sachin 3.2.1 - 05/05/08 [Added One Line below] Menus entered in Parameter 'WHHDMENU' not to be displayed if User Logs On to Full Database ******
              If Not (gs_Partition <> ctCurrPrtn And ADC.Connection.RecSeek("Select 'x' From Param Where PTyp='WHHDMENU' And PDesc='" + wo_rsusracc.FldValue("UaMnuCd") + "'")) Then
                Controls(wo_rsusracc.FldValue("uamnucd")).Visible = True
                '*** Jenny Color (10/08/05)
                If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then CreateToolBarBtn wo_rsusracc.FldValue("UaTlBarDesc"), wo_rsusracc.FldValue("uamnucd"), wi_i, wo_rsusracc.FldValue("UaMnuDesc")
                '*** Jenny Color (10/08/05)
              End If
                'Controls(wo_rsusracc.FldValue("uamnucd")).Tag = IIF(IsNull(wo_rsusracc.FldValue("uaadcopt")), "", wo_rsusracc.FldValue("uaadcopt"))
MNext:
                If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then wi_i = wi_i + 1   '*** Jenny Color (10/08/05)
                wo_rsusracc.MoveNext
            Loop
            
            '*** Jenny Color (12/08/05)
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from usraccess where uausrcd='" + UCase$(TXT_UCD.text) + "' And UaMnuForm ='LST'" + _
                              " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.text) & ",%' End")
            Do While Not (wo_rsusracc.EOF Or wo_rsusracc.BOF)
              If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then CreateToolBarBtn wo_rsusracc.FldValue("UaTlBarDesc"), wo_rsusracc.FldValue("UaMnuCd"), wi_i, wo_rsusracc.FldValue("UaMnuDesc")
              If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then wi_i = wi_i + 1
              wo_rsusracc.MoveNext
            Loop
            '*** Jenny Color (12/08/05)
            
            Set mo_CsTlBarPic = Nothing '*** Jenny Color (10/08/05)
            
            WINDOW.Visible = True
            QUIT.Visible = True
            CSC.Visible = True
            TIL.Visible = True
            TRACK.Visible = True
            GENMNU.Visible = True
            MnuTlBar.Visible = True     '*** (Jenny Color)
            Call HideDumMnus

            'Picture1.Visible = False     '*** (Bef 2.11 Color)
            StatusBar2.Visible = True
            Call InitRtn

            '*** 15/06/05 (Jenny Color)
            If wb_PicExists = True Then
              Set CMDOK.Pic = wo_Pic1
              Set CMDCAN.Pic = wo_Pic2
              Set ImgDefBtnSet.Picture = wo_PicAdc
            Else
              '*** (Bef 13/08/05)
              'Set CMDOK.Pic = go_MDIFrm.ImgDefBtnSet
              'Set CMDCAN.Pic = go_MDIFrm.ImgDefBtnSet
              '*** (Bef 13/08/05)
              
              '*** (13/08/05)
              Set CMDOK.Pic = Me.ImgDefBtnSet
              Set CMDCAN.Pic = Me.ImgDefBtnSet
              '*** (13/08/05)
              
            End If
            Set wo_Pic1 = Nothing
            Set wo_Pic2 = Nothing
            Set wo_PicAdc = Nothing
            '*** 15/06/05 (Jenny Color)

            'Call CreateToolBar    '*** Jenny Color (Bef 10/08/05)

            '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
            wMultiComp = IIF(ADC.Connection.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
            wSelfCocdNm = ADC.Connection.GetFldVal("Select HName From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'")
            '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
            
            wDbVer = ADC.Connection.GetFldVal("Select PDesc225 From Param where PTyp= 'VER' AND PMCD ='DATVER'")     '*** Jay 3.01 *** [Form Caption]
            
            Me.Caption = IIF(wMultiComp = True, wSelfCocdNm, gs_CoNm) + "  (" + Me.Caption + IIF(MWLib.IsItLaptop, " Laptop Version", "") + ", Db Svr= " + MWLib.DbSrvrNm + ", Db= " + MWLib.DbNm + "(" + IIF(TxtPartition.text = "F", "Full", "Curr") + ")" + ", Rpt Svr= " + MWLib.RepDbSrvrNm + ", Rpt Db= " + MWLib.RepDbNm + ", Ver= " + wDbVer + ", Login Comp= " + gs_CoCd + ", Usr= " + TXT_UCD.text + "," + DtShortStr + ")"
            
  
            SendKeys "%"
            
            ADC.Connection.Execute "Set Transaction Isolation Level Read UnCommitted"
'            End If
            
 


 
    If ADC.Connection.GetFldVal("Select Count(*) Cnt From Head Where HCoCd <> 'ZZZ'") = 1 Then
      gb_SingleCoCd = True
    Else
      gb_SingleCoCd = False
    End If
    
    Set mFrmTrackStatus = New EmrFrmTrackStatus

MDIEnd:
    If Not (mo_MstConn Is Nothing) Then
     Set mo_MstConn = Nothing
   End If
    Me.MousePointer = vbNormal
    MWLib.EndProcess Me
    Exit Sub
    
'****** Sachin 3.02 06/09/07 - Multi Server ******
ErrRep:
      MsgBox "Report Server Connection Could Not be Established"
      Me.MousePointer = vbNormal
      MWLib.EndProcess Me
      TxtCoCd.SetFocus
      Exit Sub
'****** Sachin 3.02 06/09/07 - Multi Server ******
    
ErrH:
    If Err.Number = 423 Then
        Resume MNext
    
    ' ***** Manali 3.8.0
    ElseIf Err.Number = 730 Then
        Resume MNext
    ' ***** Manali 3.8.0
    
    '*** Jenny Color (15/06/05)
    ElseIf Err.Number = 481 Then
        '*** Invalid Picture
        Resume Next
    '*** Jenny Color (15/06/05)
        
    Else
        MsgBox Err.Description
        
        '*** (Jen 2.14)
        Me.MousePointer = vbNormal
        MWLib.EndProcess Me
        TxtCoCd.SetFocus
        Exit Sub
        '*** (Jen 2.14)
       
        'Resume    '*** (Bef 2.14)
    End If
    Exit Sub
End Sub
Public Function CoChk(ByRef ErrMsg As String) As Boolean
  Dim wSqlStrg  As String
  
  CoChk = False
  
  Dim wRsHd As MwfLib.MDORowSet, wCompanyNm As String, wHName As String
  wCompanyNm = MWLib.CoNm
  
  If Trim(wCompanyNm) = "" Then
    ErrMsg = "Company Name in INI file is blank"
    Exit Function
  End If

  '****** Get Company Name form Head file and check if it is the same as that in the Ini file
  wSqlStrg = "Select HName From Head where HCoCd= '" + TxtCoCd + "' and HCd='" + ctSelfCmCd + "'"
      Set wRsHd = ADC.Connection.OpenRes(wSqlStrg)
  If wRsHd.RecCount <= 0 Then
    ErrMsg = "Invalid Installation Company Record not found in Head table.Company Code:" + TxtCoCd
    Exit Function
  End If
  
  If wRsHd.RecCount <> 0 Then wHName = wRsHd.FldValue("HName")
  
  If UCase(wCompanyNm) <> UCase(wHName) Then
    ErrMsg = "Company Name in Head not matching Company Name in INI file." + " Company Code: " + TxtCoCd + vbCrLf + " Name in Ini File: " + wCompanyNm + vbCrLf + " Name in Head: " + wHName
    Exit Function
  End If
  CoChk = True

End Function

'*****************  Std Code not to  be changed ****************************************'
Private Sub OldCmdOk_Click()

    '****** Sachin 3.02 [26/11/07] - Speed Optimization Version
    If Not (TxtPartition.text = "C" Or TxtPartition.text = "F") Then
        MsgBox "Please Enter Either 'C' (Current)  Or 'F'(Full Database) in Partition "
        TxtPartition.SetFocus
        Exit Sub
    End If
    '****** Sachin 3.02 [26/11/07] - Speed Optimization Version
    If Not mo_MstConn.RecSeek("Select 'x' From mSys Where  MsSrvrNm='" + ctProjectName + "' and MsLoginNm='" + cmbDbNm.text + "' and MsDbSrvrNm='" + cmbSrvrNm.text + "'") Then MsgBox "Database Record not found in Sys Table. ": Exit Sub
    
    MWLib.BeginProcess Me, "Loading " + Me.Caption
    Me.MousePointer = vbHourglass
    Dim wo_rsUsr As MwfLib.MDORowSet, wo_cn As New MwfLib.MDOConnection
    Dim wSqlStrg As String, ws_Tmp As String, wChkIntegrity As String
    '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
    Dim wMultiComp As Boolean, wSelfCocdNm As String
    '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
    Dim wDbVer As String '*** Jay 3.01 *** [Form Caption]
    '*** (Bef 27/01/07) (This works only on SQL 2005)
    ''*** (Jen speed)
    'Dim wb_HlpConn As Boolean
    'wb_HlpConn = False
    ''*** (Jen speed)
    '*** (Bef 27/01/07) (This works only on SQL 2005)
    If Not ChkInstance Then GoTo MDIEnd
    
    If Not MWLib.SetDateFormat Then GoTo MDIEnd
    If Not Ctl.SetLogInDets(TXT_UCD, ctProjectName, TxtCoCd, cmbSrvrNm.text, cmbDbNm.text) Then GoTo MDIEnd     '3.11.4

    '*** (Bef 04/02/08)
    ''*** (Jenny speed) 24/11/06
    'Call Ctl.SetConnection
    ''*** (Jenny speed) 24/11/06
    
    '*** (Bef 04/02/08)
    Set Ctl.MDIApp = Me
    If Not ADC.InitRtn Then GoTo MDIEnd
    '*** (Bef 04/02/08)
    
      ' ****** Manali 3.5.0 - Check Shifted down - to be fire after DtAndCoChk
''''    '****** Sachin 3.3.0 (01-07-08) - [16. Modules and Ini Users Check]******
''''    wChkIntegrity = CheckIntegrity(MWLib.NoOfUsers)
''''    If wChkIntegrity <> "" Then MsgBox wChkIntegrity, vbCritical, "Emperor Integrity Check Fail... ": GoTo MDIEnd
''''    '****** Sachin 3.3.0 (01-07-08) - [16. Modules and Ini Users Check]******
    
    ''*** (Jenny speed) 24/11/06
    'If Ctl.GetConnection Is Nothing Then GoTo MDIEnd
    ''*** (Jenny speed) 24/11/06
    '*** (Bef 04/02/08)

'    If MWLib.NetWrkCrdNo <> adc.Connection.GetFldVal("Select Max(Net_Address) From Master..SysProcesses Where " + _
                                            " Program_Name = 'SQLAgent - Alert Engine'") Then
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
    On Error GoTo ErrH
' ########################  Manoj #### Ver: 2.0.8 #### Date: 07/10/2004  ########################
    'Call PrnMnus
    
    '*** (Bef 27/01/07) (This works only on SQL 2005)
    ''*** (Jenny speed)
    'wb_HlpConn = True
    '
    'Call Ctl.SetHlpConnection
    'If Ctl.GetHlpConnection Is Nothing Then GoTo MDIEnd
    '
    ''On Error GoTo MDIEnd
    ''Ctl.GetHlpConnection.Execute ("Select 'x' from Head where HCoCd= '" + TxtCoCd + "' ")
    '
    ''On Error GoTo HlpEnd
    'Dim mRs_HlpTst As MDORowSet
    ''Set mRs_HlpTst = Nothing
    '''Set mRs_HlpTst = MWLib.GetHlpConn.OpenRes("Select 'x' from Head where HCoCd= '" + TxtCoCd + "' ", True)
    'Set mRs_HlpTst = Ctl.GetHlpConnection.OpenRes("Select 'x' from Head where HCoCd= '" + TxtCoCd + "' ", True)
    ''DispMsg "Count: " & CStr(mRs_HlpTst.RecCount), etError
    'wb_HlpConn = False
    'Set mRs_HlpTst = Nothing
    ''*** (Jenny speed)
    '*** (Bef 27/01/07) (This works only on SQL 2005)
    
    '****** Sachin 2.13.0 - [07-07-2006] - Skip HDD Sr Check for Laptop *****
    Dim wSysTyp As String
    wSysTyp = ADC.Connection.GetFldVal("Select IsNull(SysDb,'') From SysDb Where SysId='EMR'")
    '****** Sachin 2.13.0 - [07-07-2006] - Skip HDD Sr Check for Laptop  *****
                
    '*** (Jenny Hlp speed) 04/02/08
    Call Ctl.SetHlpConnection
    If Ctl.GetHlpConnection Is Nothing Then GoTo MDIEnd
    Set go_HlpCn = Ctl.GetHlpConnection
    '*** (Jenny Hlp speed) 04/02/08
                
    '*** (Bef 04/02/08)
    ''****** Sachin 3.02.0 [Speed Optimization] ******
    'Call Ctl.SetRepConnection
    'If Ctl.GetRepConnection.ConnState = 0 Then GoTo ErrRep
    '*** (Bef 04/02/08)
    
    gs_Partition = ctCurrPrtn      '****** Setting the Default Partition to Current
    '*** (Bef 04/02/08)
    'Set groCn = Ctl.GetRepConnection
    '*** (Bef 04/02/08)
    gbSvrChk = IIF(UCase(MWLib.DbSrvrNm) = UCase(MWLib.RepDbSrvrNm), True, False)
    
    '****** Sachin 3.02.0 - Logon Check ******
    If TxtPartition.text <> ctCurrPrtn And ADC.Connection.GetFldVal("Select PNum1 From Param Where PTyp= 'USR' And " + _
          " PMCd ='" + Trim(TXT_UCD.text) + "' And PValidYn = 'Y'") <> 2 Then
        MsgBox "User Doesn't Have Rights To Logon To Full Database", vbOKOnly, "Partition Check"
        GoTo MDIEnd
    Else
        If MWLib.DataPurged(ctProjectName) And TxtPartition.text <> ctCurrPrtn Then gs_Partition = Trim(TxtPartition.text)
    End If
    '****** Sachin 3.02.0 - Logon Check ******
        
    '****** Sachin 3.02 24-12-07 - Setting the Purge Date
    'sv.9 getting purge date from vparam
    gs_PurgeDt = GetPurgeDt(TxtCoCd.text, "JT")
    '****** Sachin 3.02.0 [Speed Optimization] ******
                
    '*** (Bef 2.13) Harddisk Serial Number Check Temporarily commented since the code is unpredictable at some places (especially on SQL Server 2005). Please verify this check thoroughly
    'If MWLib.HDDSrlNo <> "56789" And UCase(wSysTyp) <> "LTP" Then
    '  ws_Tmp = MWLib.GetSrvrHddSrlNo(MWSrvrClsID)
    '  Select Case ws_Tmp
    '    Case Is = "-1", "-2", "-3"
    '      MsgBox "Server is Not Installed Properly (" & ws_Tmp & ").", vbOKOnly, "Server Check"
    '      GoTo MDIEnd
    '    Case Else
    '      If ws_Tmp <> MWLib.HDDSrlNo Then
    '        MsgBox "Server is Not Installed Properly.", vbOKOnly, "Server Check"
    '        GoTo MDIEnd
    '      End If
    '  End Select
    'End If
    '*** (Bef 2.13) Harddisk Serial Number Check Temporarily commented since the code is unpredictable at some places (especially on SQL Server 2005). Please verify this check thoroughly
    
'    If MWLib.HDDSrlNo <> "56789" Then
'      ws_Tmp = MWLib.GetSrvrHddSrlNo(TXT_UCD.Text)
'      If ws_Tmp = "0000" Then
'        MsgBox "Server Request is Not Completed.", vbOKOnly, "Server Check"
'        GoTo MDIEnd
'      ElseIf ws_Tmp = "-1" Then
'        MsgBox "User Code is Not Present in DataBase.", vbOKOnly, "Server Check"
'        GoTo MDIEnd
'      ElseIf ws_Tmp = "-2" Then
'        GoTo MDIEnd
'      ElseIf MWLib.HDDSrlNo <> ws_Tmp Then
'        MsgBox "Server is Not Installed Properly.", vbOKOnly, "Server Check"
'        GoTo MDIEnd
'      End If
'    End If
' ###########################################  Manoj  ###########################################

'    If MWLib.HDDSrlNo <> "54321" Then
'      ws_Tmp = MWLib.GetSrvrHddSrlNo()
'      If ws_Tmp <> "0000" And MWLib.HDDSrlNo <> ws_Tmp Then
'          MsgBox "Server Component Installation is Not Proper.", vbOKOnly, "Server Check"
'          GoTo MDIEnd
'      End If
'      If ws_Tmp = "0000" Then GoTo MDIEnd
'    End If
' ###########################################  Manoj  ###########################################
    ' Zubin 213
    If Not ResolutionChk Then
      MsgBox "Screen Resolution Should Be 1024x768 Or Higher", vbOKOnly, "Screen Resolution"
      GoTo MDIEnd
    End If
    
    '*** (Jen 3.01)
    If ADC.Connection.RecSeek("Select 'x' From Param where PTyp= 'VER' and PMCd= 'DATVER' and PSCd= '' and PValidYN= 'N'") = True Then
      MsgBox "Cannot Login As Database Is In Single User Mode For Some Exclusive Updates"
      GoTo MDIEnd
    End If
    '*** (Jen 3.01)
    
    ' Zubin 213
    Call HideAllMnus
    ' Manoj 2.10.0
    wSqlStrg = "SELECT * FROM PARAM WHERE PTYP= 'USR' AND PMCD ='" + TXT_UCD.text + "' And PValidYn = 'Y'"
    ' Manoj 2.10.0
    Set wo_rsUsr = ADC.Connection.OpenRes(wSqlStrg)
    If wo_rsUsr.RecCount <> 0 Then
        If wo_rsUsr.FldValue("PValue") = TXT_PWD.text Then
            '*** Jen ***
            'If Not ValidNetId() Then MsgBox "Invalid Install": TXT_PWD.SetFocus: Exit Sub

            'If GetVersion = False Then TXT_PWD.SetFocus: Exit Sub
'            gs_Ver = GetVersion
            If gs_Ver = "" Then TXT_PWD.SetFocus: Exit Sub

            If DtAndCoChk = False Then
              If TxtCoCd.Visible = False Then TXT_PWD.SetFocus: Exit Sub
              If TxtCoCd.Visible = True Then TxtCoCd.SetFocus: Exit Sub
            End If
            
            ' ****** Manali 3.5.0 - Check to be fired after DtAndCoChk
            '****** Sachin 3.3.0 (01-07-08) - [16. Modules and Ini Users Check]******
            wChkIntegrity = CheckIntegrity(MWLib.NoOfUsers)
            If wChkIntegrity <> "" Then MsgBox wChkIntegrity, vbCritical, "Emperor Integrity Check Fail... ": GoTo MDIEnd
            '****** Sachin 3.3.0 (01-07-08) - [16. Modules and Ini Users Check]******
            
            'If FnCoCdChk = False Then TxtCoCd.SetFocus: Exit Sub
            '*** Jen ***
' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
            ' User Check
            Dim wl_UsrCnt As Long, wo_rsUL As MwfLib.MDORowSet
            ' No of Allowed User Check
            '*** Jay 3.1.0_1 [UlCoCd Added]
            'wl_UsrCnt = CLng(ADC.Connection.GetFldVal("Select isNull(Count(*), 0) from UsrLogin"))
            
            '****** Manali 3.8.0 - Position Of Following Check - Shifted before Ini  user check
            gs_HostNm = UCase$(GetHostNm)
            
            Dim wi_delusr As Integer
            If ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' /* and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' */ and UlHostNm='" + gs_HostNm + "'") Then
               wi_delusr = MsgBox("User '" & UCase$(TXT_UCD.text) & "' Already Logged In " + UCase$(Trim$(TxtCoCd.text)) + " Company On This Machine. Delete User ?", vbYesNo, "User Check")
               If wi_delusr = vbYes Then
                 ADC.Connection.Execute ("Delete From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' /* and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' */ And UlHostNm='" + gs_HostNm + "' ")
                 'GoTo MDIAddUsrRec
               Else
                 GoTo MDIEnd
               End If
            End If
            '****** Manali 3.8.0 - Position Of Following Check - Shifted before Ini  user check
            
            wl_UsrCnt = CLng(ADC.Connection.GetFldVal("Select Count(Distinct(UlUsrCd)) from UsrLogin"))
            '*** Jay 3.1.0_1 [UlCoCd Added]
            '****** Sachin 2.13.0 - [08-07-2006] - If Laptop, Restrict Users to 2 ******
            If wl_UsrCnt >= IIF(UCase(wSysTyp) <> "LTP", MWLib.NoOfUsers, 2) And MWLib.NoOfUsers <> 321 Then
              MsgBox "No of Users Allowed (" & CStr(IIF(UCase(wSysTyp) <> "LTP", MWLib.NoOfUsers, 2)) & ") are Already Logged On", vbOKOnly, "User Count Check"
              GoTo MDIEnd
            Else
               '*** Jay 3.1.0_1 [UlCoCd Added]
'               gs_HostNm = UCase$(GetHostNm)
               'User Already Logon Check
               '*** BEF 3.1.0_1
               'If ADC.Connection.RecSeek("Select UlUsrCd From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "'") Then
               '  MsgBox "User '" & UCase$(TXT_UCD.text) & "' Already Logged On", vbOKOnly, "User Check"
               '  GoTo MDIEnd
               '*** BEF 3.1.0_1
               
''              Manali 3.8.0 - Position Of Following Check Changed - Shifted before Ini  user check
'               '****** Sachin 3.5.0 [11/11/08] - IdNo Check In UsrLogin ******
'               Dim wi_delusr As Integer
'               If adc.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' and UlHostNm='" + gs_HostNm + "'") Then
'                  wi_delusr = MsgBox("User '" & UCase$(TXT_UCD.text) & "' Already Logged In " + UCase$(Trim$(TxtCoCd.text)) + " Company On This Machine. Delete User ?", vbYesNo, "User Check")
'                  If wi_delusr = vbYes Then
'                    adc.Connection.Execute ("Delete From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' And UlHostNm='" + gs_HostNm + "' ")
'                    GoTo MDIAddUsrRec
'                  Else
'                    GoTo MDIEnd
'                  End If
'               '****** Sachin 3.5.0 [11/11/08] - IdNo Check In UsrLogin ******
''              Manali 3.8.0 - Position Of Following Check Changed - Shifted before Ini  user check

               If wi_delusr = vbYes Then
                 GoTo MDIAddUsrRec
               ElseIf ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "'") Then
                  MsgBox "User '" & UCase$(TXT_UCD.text) & "' Already Logged In " + UCase$(Trim$(TxtCoCd.text)) + " Company", vbOKOnly, "User Check"
                  GoTo MDIEnd
               ElseIf ADC.Connection.GetFldVal("Select Count('x') From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "'") >= 2 Then
                  MsgBox "User '" & UCase$(TXT_UCD.text) & "' Already Logged Into Two Companies.", vbOKOnly, "User Check"
                  GoTo MDIEnd
               ElseIf ADC.Connection.RecSeek("Select 'x' where ((Not (Select Count(*) from UsrLogin Where UlUsrCd = '" + UCase$(Trim$(TXT_UCD.text)) + "' and " + _
                      "UlHostNm = '" + gs_HostNm + "') > 0 )and " + _
                      "(Select Count('x') from UsrLogin Where UlUsrCd = '" + UCase$(Trim$(TXT_UCD.text)) + "') > 0) ") Then
                  MsgBox "User '" & UCase$(TXT_UCD.text) & "' Cannot Log On Another Machine.", vbOKOnly, "User Check"
                  GoTo MDIEnd
               '*** Jay 3.1.0_1 [UlCoCd Added]
               Else
                'gs_HostNm = UCase$(GetHostNm)      '*** BEF 3.1.0_1
                'Insert Record for new Login
MDIAddUsrRec:
                Set wo_rsUL = ADC.Connection.OpenResultset("Select * from UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "'")
                If Not (wo_rsUL Is Nothing) Then
                  With wo_rsUL
                    .Addnew
                    .FldValue("UlUsrCd") = UCase$(Trim$(TXT_UCD.text))
                    .FldValue("UlCoCd") = UCase$(Trim$(TxtCoCd.text))
                    .FldValue("UlHostNm") = gs_HostNm
                    .FldValue("UlDt") = ADC.Connection.SrvrDate
                    .FldValue("UlTime") = ADC.Connection.SrvrTime
                    .Update
                  End With
                End If
              End If
            End If
'' ###########################################  Manoj  ###########################################

            '****** Sachin 3.5.0 [11/11/08] - Setting the UsrLogin IdNo In Global Variable ******
            gs_UlIdNo = ADC.Connection.GetFldVal("Select UlIdNo From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.text)) + "' " + _
                                                " and UlHostNm='" + gs_HostNm + "'")
            '****** Sachin 3.5.0 ******

            '*** Jenny Color
            Dim ws_ClrCd As String
            Dim wo_Pic1 As Picture, wo_Pic2 As Picture, wo_PicAdc As Picture, wb_PicExists As Boolean
            
            '*** (Jen 2.13)
            wb_ApplyClr = False
            
            ws_ClrCd = ADC.Connection.GetFldVal("Select PValue2 from Param where PTyp= 'USR' and PMCd= '" + TXT_UCD + "'")
            Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
            
            If Not (GetRecClrSet.RecCount > 0) Then
              ws_ClrCd = ADC.Connection.GetFldVal("Select PValue From Param where PTyp= 'DEFCLR'")
              If ws_ClrCd <> "" Then Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
            End If
            
            If GetRecClrSet.RecCount > 0 Then wb_ApplyClr = True
            
            If wb_ApplyClr = True Then
            '*** (Jen 2.13)
            
              ws_ClrCd = ADC.Connection.GetFldVal("Select PValue2 from Param where PTyp= 'USR' and PMCd= '" + TXT_UCD + "'")
              Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
              
              
              If Not (GetRecClrSet.RecCount > 0) Then
                ws_ClrCd = ADC.Connection.GetFldVal("Select PValue From Param where PTyp= 'DEFCLR'")
                Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
              End If
              
              Call SetClrPth("")
              If InStrRev(App.Path, "\") > 1 Then Call SetClrPth(Mid(App.Path, 1, InStrRev(App.Path, "\") - 1))
              If GetClrPth <> "" Then Call SetClrPth(GetClrPth & "\MwClrSet\")
              Picture1.Visible = False   '*** 15/06/05
              
              wb_PicExists = True
              
              If Trim(GetRecClrSet!CsADCBtnPic) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsADCBtnPic) = False Or Trim(GetRecClrSet!CsBtnPicSet1) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsBtnPicSet1) = False Or Trim(GetRecClrSet!CsBtnPicSet2) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsBtnPicSet2) = False Or Trim(GetRecClrSet!CsFormPic) = "" Or IsValidPath(GetClrPth & GetRecClrSet!CsFormPic) = False Then
                wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_PicAdc = LoadPicture("")    '*** This will assign 0 to woPic
                Set wo_PicAdc = LoadPicture(GetClrPth & GetRecClrSet!CsADCBtnPic)
                If (wo_PicAdc Is Nothing) Or wo_PicAdc = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_Pic1 = LoadPicture("")
                Set wo_Pic1 = LoadPicture(GetClrPth & GetRecClrSet!CsBtnPicSet1)
                If (wo_Pic1 Is Nothing) Or wo_Pic1 = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set wo_Pic2 = LoadPicture("")
                Set wo_Pic2 = LoadPicture(GetClrPth & GetRecClrSet!CsBtnPicSet2)
                If (wo_Pic2 Is Nothing) Or wo_Pic2 = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = True Then
                Set Picture1.Picture = LoadPicture("")
                Set Picture1.Picture = LoadPicture(GetClrPth & GetRecClrSet!CsFormPic)
                If (Picture1.Picture Is Nothing) Or Picture1.Picture = 0 Then wb_PicExists = False
              End If
              If wb_PicExists = False Then
                ws_ClrCd = ctDefClrCd
                Call SetRecClrSet(ADC.Connection.OpenRes("Select * from ClrSet where CsCd= '" + ws_ClrCd + "' "))
              End If
            
            '*** (Jen 2.13)
            Else
              RestoreColors
              Picture1.Visible = False
              Me.Picture = LoadPicture("")
              Call SetRecClrSet(Nothing)
              Set mo_MDIPic = Nothing
              Set mo_CsTlBarPic = Nothing
            End If
            '*** (Jen 2.13)
            
            SetMDISysColorsAfterLogin
            '*** (Jen 2.13)
            If Not (GetRecClrSet Is Nothing) Then
            '*** (Jen 2.13)
              If GetRecClrSet.RecCount > 0 Then
                GetRecClrSet.MoveFirst
                If Trim(GetRecClrSet!CsMDIPic) <> "" And IsValidPath(GetClrPth & GetRecClrSet!CsMDIPic) = True Then
                  Set mo_MDIPic = LoadPicture(GetClrPth & GetRecClrSet!CsMDIPic)
                End If
                If Not (mo_MDIPic Is Nothing) Then Set PictureClip1.Picture = mo_MDIPic
                SetMDIPic
                Me.BackColor = Val(GetRecClrSet!CsLblForClr)
              End If
            
            '*** (Jen 2.13)
            End If
            '*** (Jen 2.13)
            '*** Jenny Color

            gs_UsrCd = TXT_UCD
            WindowState = 2

            ' Zubin 212 (27/02/06)
            Dim wo_RsModules As MwfLib.MDORowSet
            Dim ws_Mod As String
            '****** (Jen 3.01) added HCustomiseOrd ******
            '*** Jay 3.01 *** [HLclInvYn, HFgTagMod added]
            ' **** Manali 3.4.1 - [HDsgCatYn added]
            ' **** Manali 3.5.0 - 14/11/08 - [HLclInvYn, HFgTagMod Removed from following list]
            Set wo_RsModules = ADC.Connection.OpenRes("select HCoCd, HPPCYN, HCtbMod, HFgQuotYN, HLooseInv, " + _
                               "HConsolidatedInv, hAutomateReturns, /* HLclInvYn, HFgTagMod,*/ HCustomiseOrd, HTravelExlYN, HDsgCatYN from Head " + _
                               "where HCoCd ='" + UCase(TxtCoCd.text) + "' and HCd= '" + ctSelfCmCd + "'")
            If UCase(wo_RsModules.FldValue("HPPCYN")) = "N" Then ws_Mod = ws_Mod + "'PPC','PPCCAP','PPCDSGLD','PPCORDPLN','REPPPC','REPDSGLD','REPCNTRLCPY','REPORDPLN'"
            If UCase(wo_RsModules.FldValue("HCtbMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'CTB','CTBIN','CTBINALLBAGS','CTBINDSG','REPINCTB','REPDETINVCTB','REPAVGINVCTB','REPINVCTBCAT','REPCTBPCKLST'"
            If UCase(wo_RsModules.FldValue("HFgQuotYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'QB'"
            '*** (Jen 2.12 Patch 3)
            If UCase(wo_RsModules.FldValue("HLooseInv")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepInLoose'"   '"'RepInLoose', 'RepInvLoose', 'RepShpBillLoose', 'RepGrFormLoose', 'RepPckLstLoose'"
            '*** (Jen 2.12 Patch 3)
            
            '*** (Jen 2.13)
            '****** (Jen 3.01) added ('RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti') ******
            '*** Jay 3.1.1 [added RepGrFormFrntNewMul, RepGRFormBackMulti]
            If UCase(wo_RsModules.FldValue("HConsolidatedInv")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepGrFormFrntMulti', 'RepShpBillMulti', 'RepGSPFORMMulti', 'RepGSPDetLetMulti', 'RepGSPCertMulti', 'RepDeclnLetMulti', 'RepGrFormFrntNewMul', 'RepGRFormBackMulti'"
            If UCase(wo_RsModules.FldValue("HAutomateReturns")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RETMEM'"
            '*** (Jen 2.13)
            
            ' **** Manali 3.5.0 - 14/11/08 - [HLclInvYn, HFgTagMod Removed from Module list]
            '*** Jay 3.01 *** [Local Inv]
            'If UCase(wo_RsModules.FldValue("HLclInvYn")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepLclInv'"
            '*** Jay 3.01 *** [FgTag Module]
            'If UCase(wo_RsModules.FldValue("HFgTagMod")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepTagPrnBagFg', 'RepTagPrnFg'"
            ' **** Manali 3.5.0 - 14/11/08 - [HLclInvYn, HFgTagMod Removed from Module list]
            
            '*** (Jen 3.01)
            If UCase(wo_RsModules.FldValue("HCustomiseOrd")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'IdeDfn', 'IdeUsrRight'"
            '*** (Jen 3.01)
            
            '****** Sachin 3.3.0 - Hiding Travel Excel Module ******
            If UCase(wo_RsModules.FldValue("HTravelExlYN")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'RepTravExcel'"
            '****** Sachin 3.3.0 - Hiding Travel Excel Module ******
            
            ' ***** Manali - 3.4.1 - 22/10/08 - Desgin Catalog Module
            ' ***** Manali 3.6.0 - 30/09/09 - Catalogue Entry removed from Module Check
            If UCase(wo_RsModules.FldValue("HDsgCatYn")) = "N" Then ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'DsgPrm'/*, 'DsgPrmCat'*/ "
            ' ***** Manali - 3.4.1 - 22/10/08 - Desgin Catalog Module
            
            ' Zubin 212P2
            If Not ADC.Connection.RecSeek("select 'x' from SysDb where SysId= 'CRM' ") Then
              ws_Mod = ws_Mod + IIF(ws_Mod <> "", ",", "") + "'SALORDACK'"
            End If
            ' Zubin 212P2
                        
            ws_Mod = IIF(ws_Mod <> "", " and UaMnuCd Not In (" + ws_Mod + ")", "")
            Set wo_RsModules = Nothing
            ' Zubin 212 (27/02/06)
            
            Dim wo_rsusracc As MwfLib.MDORowSet
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
            ' Manoj 2.10.0
            Dim wsCoCd As String
            ' Zubin 212 (27/02/06) ( + ws_Mod ADDED)
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from UsrAccess where uausrcd='" + UCase$(TXT_UCD.text) + "' And UaMnuForm = '" + ms_MnuForm + "'" + _
                      " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.text) & ",%' End" + ws_Mod)
            ' Manoj 2.10.0
' ###########################################  Manoj  ###########################################
            
            '*** Jenny Color (10/08/05)
            Call CreateToolBar
            Dim wi_i As Integer
            wi_i = 0
            '*** Jenny Color (10/08/05)
         
            Do While Not (wo_rsusracc.EOF Or wo_rsusracc.BOF)
              '****** Sachin 3.2.1 - 05/05/08 [Added One Line below] Menus entered in Parameter 'WHHDMENU' not to be displayed if User Logs On to Full Database ******
              If Not (gs_Partition <> ctCurrPrtn And ADC.Connection.RecSeek("Select 'x' From Param Where PTyp='WHHDMENU' And PDesc='" + wo_rsusracc.FldValue("UaMnuCd") + "'")) Then
                  Controls(wo_rsusracc.FldValue("uamnucd")).Visible = True
                  '*** Jenny Color (10/08/05)
                  If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then CreateToolBarBtn wo_rsusracc.FldValue("UaTlBarDesc"), wo_rsusracc.FldValue("uamnucd"), wi_i, wo_rsusracc.FldValue("UaMnuDesc")
                  '*** Jenny Color (10/08/05)
              End If
                'Controls(wo_rsusracc.FldValue("uamnucd")).Tag = IIF(IsNull(wo_rsusracc.FldValue("uaadcopt")), "", wo_rsusracc.FldValue("uaadcopt"))
MNext:
                If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then wi_i = wi_i + 1   '*** Jenny Color (10/08/05)
                wo_rsusracc.MoveNext
            Loop
            
            '*** Jenny Color (12/08/05)
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from usraccess where uausrcd='" + UCase$(TXT_UCD.text) + "' And UaMnuForm ='LST'" + _
                              " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.text) & ",%' End")
            Do While Not (wo_rsusracc.EOF Or wo_rsusracc.BOF)
              If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then CreateToolBarBtn wo_rsusracc.FldValue("UaTlBarDesc"), wo_rsusracc.FldValue("UaMnuCd"), wi_i, wo_rsusracc.FldValue("UaMnuDesc")
              If Trim(wo_rsusracc.FldValue("UaTlBarDesc")) <> "" Then wi_i = wi_i + 1
              wo_rsusracc.MoveNext
            Loop
            '*** Jenny Color (12/08/05)
            
            Set mo_CsTlBarPic = Nothing '*** Jenny Color (10/08/05)
            
            WINDOW.Visible = True
            QUIT.Visible = True
            CSC.Visible = True
            TIL.Visible = True
            GENMNU.Visible = True
            MnuTlBar.Visible = True     '*** (Jenny Color)
            Call HideDumMnus

            'Picture1.Visible = False     '*** (Bef 2.11 Color)
            StatusBar2.Visible = True
            Call InitRtn

            '*** 15/06/05 (Jenny Color)
            If wb_PicExists = True Then
              Set CMDOK.Pic = wo_Pic1
              Set CMDCAN.Pic = wo_Pic2
              Set ImgDefBtnSet.Picture = wo_PicAdc
            Else
              '*** (Bef 13/08/05)
              'Set CMDOK.Pic = go_MDIFrm.ImgDefBtnSet
              'Set CMDCAN.Pic = go_MDIFrm.ImgDefBtnSet
              '*** (Bef 13/08/05)
              
              '*** (13/08/05)
              Set CMDOK.Pic = Me.ImgDefBtnSet
              Set CMDCAN.Pic = Me.ImgDefBtnSet
              '*** (13/08/05)
              
            End If
            Set wo_Pic1 = Nothing
            Set wo_Pic2 = Nothing
            Set wo_PicAdc = Nothing
            '*** 15/06/05 (Jenny Color)

            'Call CreateToolBar    '*** Jenny Color (Bef 10/08/05)

            '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
            wMultiComp = IIF(ADC.Connection.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
            wSelfCocdNm = ADC.Connection.GetFldVal("Select HName From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'")
            '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
            
            wDbVer = ADC.Connection.GetFldVal("Select PDesc225 From Param where PTyp= 'VER' AND PMCD ='DATVER'")     '*** Jay 3.01 *** [Form Caption]
            
            '(Old Code) Me.Caption = gs_CoNm + "  (" + Me.Caption + ")"
            '****** Sachin 2.13.0 - [09-07-2006] - Indication of Laptop Version in Caption ******
            '****** Sachin 2.14.0 - [24-11-2006] - If Multi-Company is enabled, Self Company Name is Displayed ******
            '*** Jay 3.01 *** [Form Caption]
            'Me.Caption = IIF(wMultiComp = True, wSelfCocdNm, gs_CoNm) + "  (" + Me.Caption + IIF(UCase(wSysTyp) = "LTP", " Laptop Version", "") + ", Db Srvr= " + MWLib.DbSrvrNm + ", Db= " + MWLib.DbNm + ", Usr= " + TXT_UCD.text + ")" '*** Jen
            Me.Caption = IIF(wMultiComp = True, wSelfCocdNm, gs_CoNm) + "  (" + Me.Caption + IIF(UCase(wSysTyp) = "LTP", " Laptop Version", "") + ", Db Svr= " + MWLib.DbSrvrNm + ", Db= " + MWLib.DbNm + "(" + IIF(TxtPartition.text = "F", "Full", "Curr") + ")" + ", Rpt Svr= " + MWLib.RepDbSrvrNm + ", Rpt Db= " + MWLib.RepDbNm + ", Ver= " + wDbVer + ", Login Comp= " + gs_CoCd + ", Usr= " + TXT_UCD.text + ")"
            '*** Jay 3.01 *** [Form Caption]
            'Call ADC.CloseConn

            SendKeys "%"
            
            '****** Sachin 3.2.0 - Serializable to be set if Purge Not done or Logged on to Full Database
            If ((MWLib.DataPurged(ctProjectName) = True And gs_Partition <> ctCurrPrtn) Or _
                (MWLib.DataPurged(ctProjectName) = False And gs_Partition = ctCurrPrtn)) Then
                ADC.Connection.Execute "Set Transaction Isolation Level Serializable"
            Else
                ADC.Connection.Execute "Set Transaction Isolation Level Read UnCommitted"
            End If
            
            'ADC.Connection.Execute "Set Transaction Isolation Level Read UnCommitted"
' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
            If ADC.Connection.RecSeek("Select 1 from UsrMail Where UmUsrCd ='" + gs_UsrCd + "' And UmMailAct ='Y'") Then _
              Call MailLogin
            Call SetMailMenu
' ###########################################  Manoj  ###########################################
        Else
            MsgBox ("Invalid Password")
            TXT_PWD.SetFocus
        End If
    Else
        MsgBox ("Invalid User code")
        TXT_UCD.SetFocus
    End If

' ########################  Manoj #### Ver: 2.0.9 #### Date: 22/11/2004  ########################
    If ADC.Connection.GetFldVal("Select Count(*) Cnt From Head Where HCoCd <> 'ZZZ'") = 1 Then
      gb_SingleCoCd = True
    Else
      gb_SingleCoCd = False
    End If
' ###########################################  Manoj  ###########################################

  '***Geeta****Crm
    If ADC.Connection.GetFldVal("Select count(*) From SysDb where SysId='CRM'") = 1 Then
      gb_CrmDb = True
    Else
      gb_CrmDb = False
    End If

    '*** (Bef 04/02/08)
    ''*** (Jenny speed 24/11/06)
    'Set goCn = Ctl.GetConnection
    ''*** (Jenny speed 24/11/06)
    '*** (Bef 04/02/08)


MDIEnd:
    Me.MousePointer = vbNormal
    MWLib.EndProcess Me
    Exit Sub
    
'****** Sachin 3.02 06/09/07 - Multi Server ******
ErrRep:
      MsgBox "Report Server Connection Could Not be Established"
      Me.MousePointer = vbNormal
      MWLib.EndProcess Me
      TxtCoCd.SetFocus
      Exit Sub
'****** Sachin 3.02 06/09/07 - Multi Server ******
    
ErrH:

    '*** (Bef 27/01/07) (This works only on SQL 2005)
    ''*** Jen speed)
    'If wb_HlpConn = True Then
    '  Me.MousePointer = vbNormal
    '  MWLib.EndProcess Me
    '
    '  'MsgBox CStr(Err.Number) & ": " & Err.Description
    '  wb_HlpConn = False
    '  MsgBox "Connection Could Not Be Established"
    '  TxtCoCd.SetFocus
    '  Exit Sub
    'End If
    ''*** Jen speed)
    '*** (Bef 27/01/07) (This works only on SQL 2005)
    
    If Err.Number = 423 Then
        Resume MNext
        
    ' ***** Manali 3.8.0
    ElseIf Err.Number = 730 Then
        Resume MNext
    ' ***** Manali 3.8.0
    
    '*** Jenny Color (15/06/05)
    ElseIf Err.Number = 481 Then
        '*** Invalid Picture
        Resume Next
    '*** Jenny Color (15/06/05)
        
    Else
        MsgBox Err.Description
        
        '*** (Jen 2.14)
        Me.MousePointer = vbNormal
        MWLib.EndProcess Me
        TxtCoCd.SetFocus
        Exit Sub
        '*** (Jen 2.14)
       
        'Resume    '*** (Bef 2.14)
    End If
    Exit Sub
End Sub
Public Sub Ctl_HandleError(ByVal pv_ErrMsg As String, ByVal pv_ErrorType As MwfCtl.en_ErrorType)
  If UCase(Trim(pv_ErrMsg)) = UCase(Trim("Out of Memory")) Then pv_ErrMsg = Trim(pv_ErrMsg) + ". Please Reduce The Scope."
  If g_InProcessingMode And (pv_ErrorType = etInfo Or pv_ErrorType = etWarning) Then
    'DONOTHING
  Else
    StatusBar2.Panels("ErrMsg").text = pv_ErrMsg
  End If
  If pv_ErrorType = etInfo And g_InProcessingMode Then
   If IsItUpdateEmrProcessMsg(pv_ErrMsg) Then
    StatusBar2.Panels("ErrMsg").text = pv_ErrMsg
   End If
  End If
  '*** (Jen 2.13)
  If pv_ErrorType = etError Then
   gs_ErrMsg = pv_ErrMsg


  End If

  If pv_ErrorType = etError Then MWLib.GenError pv_ErrMsg
  If pv_ErrorType = etWarning Then
    MWLib.GenError pv_ErrMsg, True
  End If
End Sub

'*** (Bef 12/08/05)
'Private Sub ShowForm(ByVal Frm As Form, MenuCd As String)
'  Dim wSqlStrg  As String, wf_ExecGot As Boolean
'  On Error GoTo errh
'
'  '*** Jenny Color
''  gs_MnuCd = MenuCd   '*** 13/06/05
''  If GetRecClrSet.RecCount > 0 Then
''    GetRecClrSet.MoveFirst
''    '*** For Setting System Colors (For Tab Page)
''    SetTabSysColors Val(GetRecClrSet!CsTabPgClr), Val(GetRecClrSet!CsTabForClr), Val(GetRecClrSet!CsTabBorderClr), Val(GetRecClrSet!CsTabShadowClr), Val(GetRecClrSet!CsTabHighLightClr)
''    '*** For Setting System Colors (For Tab Page)
''  End If
''  If (Frm Is Nothing) Then Exit Sub
'  '*** Jenny Color
'
'
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
'' Check Session is Working or Not
'    If Not ChkLogin(ADC.Connection) Then End
'' ###########################################  Manoj  ###########################################
'    MWLib.BeginProcess Me, "Loading " + Controls(MenuCd).Caption
'    wf_ExecGot = False
'    If Not (Me.ActiveForm Is Nothing) Then
'        Me.ActiveForm.ADC.ExecGot = True
'        wf_ExecGot = True
'    End If
'    Load Frm
'    takecloseoff Frm.hwnd
'    If wf_ExecGot Then Frm.ADC.ExecGot = True
'    If Frm.WindowState = 0 Then
'      Frm.Left = 0
'      Frm.Top = 0
'    End If
'    Frm.ADC.MenuCd = MenuCd
'    Frm.ADC.UsrCd = TXT_UCD
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
'' Set HostName  as ADC.HostNm for User Check
'    Frm.ADC.HostNm = gs_HostNm
'' Set MDI Form as MenuForm for Loading Form
'    Set Frm.ADC.MnuForm = Me
'' ###########################################  Manoj  ###########################################
'
'    '*** Jenny Color (02/04/05)
'    Frm.ADC.ToolBarCtrlName = "CmdCtlToolBar"
'    Frm.ADC.ToolBarCtrlIndx = mo_TlBarColl.Item(MenuCd)
'    '*** Jenny Color (02/04/05)
'
'    If Not Frm.ADC.InitRtn Then End
'    Set Frm = Nothing
'    MWLib.EndProcess Me
'    Exit Sub
'errh:
'    If Err.Number = 40002 Then
'        Resume Next
'
'    '*** Jenny Color (02/04/05)
'    ElseIf Err.Number = 9 Or Err.Number = 5 Then
'        '*** 9: If the reqd Menucode does not exist in the collection then this error is generated
'        Frm.ADC.ToolBarCtrlName = ""
'        Resume Next
'    '*** Jenny Color (02/04/05)
'
'    Else
'        MsgBox Err.Description
'    End If
'End Sub
'*** (Bef 12/08/05)

Private Sub MDIForm_Load()
    '*** Jenny Color
    'Set go_MDIFrm = Me    '(16/05/05) (Bef 13/08/05)
    Call SetMDIFrm(Me)    '(13/08/05)
    '*** Jenny Color
    
    takecloseoff hwnd
    Set MWLib = New MWCLS_GEN

    '*** Jenny
    TxtCoCd.text = MWLib.SingleCoCd(ctProjectName)
    If TxtCoCd.text <> "" Then
      LblCoCd.Visible = False
      LblCoCd.Enabled = False
      TxtCoCd.Visible = False
    Else
      LblCoCd.Visible = True
      LblCoCd.Enabled = True
      TxtCoCd.Visible = True
    End If
    
    '*** Jenny
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
    'ms_MnuForm = "JMD"
    ms_MnuForm = "MDI"
' ###########################################  Manoj  ###########################################


    '*** Jenny Color
    wb_ApplyClr = False   '*** (Jen 2.13)
    
    TlBar.Visible = False
    Picture1.Picture = Me.Picture
    
    '****** (For Setting System Colors)
      '*** Store handle to this form's window.
      '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
      'gHW = Me.hWnd
      '##$$ (Bef 09/08/05) This is necessary for Hook & UnHook to work. Uncomment it whenever required
  
      '*** Call procedure to begin capturing messages for this window.
      '##$$ Hook      '*** Imp Note: commented temporarily because while debugging, this option is a hindrance for the programmer
  
      '*** Call procedure to save the original color settings.
      SaveOriginalColors
  
      '*** Call procedure to set the new colors as the ones in use.
      SetMDISysColorsBefLogin
      
    '****** (For Setting System Colors)
    '*** Jenny Color
    
    '****** Sachin 3.02.0
   ' TxtPartition.Visible = MWLib.DataPurged(ctProjectName)
   ' LblPartition.Visible = MWLib.DataPurged(ctProjectName)
    '****** Sachin 3.02.0
    If Not MWLib.SetDateFormat Then GoTo MDIEnd
    
    MWLib.SetLoginParameters "Mw", "EmrSQLMw007*99", "", ""
    'MWLib.SetLoginParameters "sa", "sa123", "", ""
    'Ctl.SetLoginParameters "UsrApp", "UsrAppPwd", "AppRole", "AppRolePwd"
    Ctl.SetLoginParameters "Mw", "EmrSQLMw007*99", "", ""
    'Ctl.SetLoginParameters "sa", "sa123", "", ""
    
    Set mo_MstConn = MWLib.GetMwMstConnection(ctProjectName)
    
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If

 '    Call FillSrvrCmb
'     Exit Sub
    
'    Set mo_MstConn = MWLib.GetMwMstConn(ctProjectName)
 '   If mo_MstConn Is Nothing Then
  '      DispMsg "Unable to Connect to Master Server", etError
   '     Exit Sub
    ' End If

    Call FillSrvrCmb        '3.11.4
    
    Exit Sub
MDIEnd:
    Me.MousePointer = vbNormal
    MWLib.EndProcess Me
    Unload Me
    Exit Sub
   
End Sub
Private Sub MDIForm_Resize()
  '*** Jenny Color (For Setting System Colors)
  If wb_ApplyClr = False Then Exit Sub  '*** (Jen 2.13)
  
  If Me.WindowState <> vbMinimized Then
    If gs_CoCd = "" Then
      SetMDISysColorsBefLogin
    Else
      SetMDISysColorsAfterLogin
    End If
    'SetNewColors
    SetMDIPic
  Else
     RestoreColors
  End If
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'*** Jenny Color (For Setting System Colors)
  '*** Call procedure to restore colors to their original.
  RestoreColors
  
  '*** (09/08/05)
  '*** Jenny Color
  'Set go_RecClr = Nothing
  Call SetRecClrSet(Nothing)
  'Set go_MDIFrm = Nothing
  Call SetMDIFrm(Nothing)
  Set mo_MDIPic = Nothing
  Set mo_CsTlBarPic = Nothing
  Set mo_TlBarColl = Nothing
  '*** Jenny Color


  Set MWLib = Nothing
  Set HlpList = Nothing
  Set gColTgFldTyp = Nothing
  Set mwhlp = Nothing
  '*** (09/08/05)

  
  '*** Call subprocedure to cease hooking into messages.
  '##$$ Unhook    '*** Imp Note: commented temporarily because while debugging, this option is a hindrance for the programmer
'*** Jenny Color (For Setting System Colors)
End Sub
Private Sub quit_Click()
    Unload mFrmTrackStatus
    Set mFrmTrackStatus = Nothing
    If Not (Me.ActiveForm Is Nothing) Then
       Call Ctl_HandleError("Close all Child forms before quitting", etError)
    Else
       Set Ctl.ActScrn = Nothing
       Set Ctl.MDIApp = Nothing
       
       '*** (Jenny speed)
       'Call Ctl.CloseConnections
       '*** (Jenny speed)
       
       '*** (Jen 04/02/08)
       Set MWLib = Nothing   '*** (Bef 24/11/06)
       '*** (Jen 04/02/08)

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' If Application is Closed Remove the Login Entry
        Dim wo_rsUL As MwfLib.MDORowSet
        '*** Jay 3.1.0_1 [UlCoCd Added]
        '6.1-49 UlSysCd='JEMR' added
        Set wo_rsUL = ADC.Connection.OpenResultset("Select * from UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.text)) + "' " + _
                            "and UlCoCd='" + UCase$(Trim$(TxtCoCd)) + "' And UlHostNm = '" + gs_HostNm + "' and UlSysCd='JEMR'")
        If Not (wo_rsUL Is Nothing) Then
          With wo_rsUL
            If .RecCount > 0 Then
              .MoveFirst
              .Delete
            End If
          End With
        End If
' ###########################################  Manoj  ###########################################

      ''*** (Jenny speed)   24/11/06
      Call Ctl.CloseConnections
      '*** (Jen 04/02/08)
      Set go_HlpCn = Nothing
      'Set mo_HlpConn = Nothing
      '*** (Jen 04/02/08)
      
      '*** (Bef 04/02/08)
      'Set goCn = Nothing
      'Set MWLib = Nothing
      '*** (Bef 04/02/08)
      ''*** (Jenny speed)   24/11/06

' ########################  Manoj #### Ver: 2.11.0 #### Date: 26/05/2005  ########################
       Call LogOffSes
' ###########################################  Manoj  ###########################################
       Unload Me
    End If
End Sub
Private Sub til_Click()
    EmrFrmJMDI.Arrange vbTileVertical
End Sub
Public Sub takecloseoff(handle As Long)
    Dim SysMenHandle As Long, RetVal As Long
    SysMenHandle = GetSystemMenu(handle, 0)
    RetVal = RemoveMenu(SysMenHandle, 6, MF_BYPOSITION)
    RetVal = RemoveMenu(SysMenHandle, 5, MF_BYPOSITION)
End Sub
Private Sub Track_click()
    mFrmTrackStatus.Show 0, Me
End Sub
Private Sub GenMnu_click()
  On Error GoTo GenMnuErr
    Dim rs_UsrAcc As MwfLib.MDORowSet
    Dim wo_Ctl As Control
    MWLib.BeginProcess Me, "Menu Generation Is In Progress"
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
    ADC.Connection.Execute ("DELETE FROM USRACCess WHERE UAUSRCD= '" + ctGenMenuUsr + "' And UAMNUFORM = '" + ms_MnuForm + "'")
    Set rs_UsrAcc = ADC.Connection.OpenResultset("select * from usraccess where uausrcd= '" + ctGenMenuUsr + "' And UAMNUFORM = '" + ms_MnuForm + "'")
' ###########################################  Manoj  ###########################################
    For Each wo_Ctl In Controls
        If TypeOf wo_Ctl Is Menu Then
'??? consider speed
'****** (Jen 3.01 added IdeDfn, RepDeclnLetMulti) ******
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
            If UCase$(wo_Ctl.Caption) <> "DUM" And UCase$(wo_Ctl.Name) <> "WINDOW" And UCase$(wo_Ctl.Name) <> "CSC" And _
              UCase$(wo_Ctl.Name) <> "TIL" And UCase$(wo_Ctl.Name) <> "GENMNU" And UCase$(wo_Ctl.Name) <> "QUIT" _
              And UCase$(wo_Ctl.Name) <> UCase$("IdeDfn") And UCase$(wo_Ctl.Name) <> UCase$("RepDeclnLetMulti") Then
' ###########################################  Manoj  ###########################################
                'Debug.Print wo_Ctl.Name
                rs_UsrAcc.Addnew
                rs_UsrAcc!UAUSRCD = ctGenMenuUsr
                rs_UsrAcc!UaMnuCd = UCase(Trim(wo_Ctl.Name))
                rs_UsrAcc!UaMnuDesc = Trim(wo_Ctl.Caption)
                rs_UsrAcc!UaAdcOpt = ""
' ########################  Manoj #### Ver: 2.0.6 #### Date: 03/06/2004  ########################
                rs_UsrAcc!UaMnuForm = ms_MnuForm
' ###########################################  Manoj  ###########################################
' ########################  Manoj #### Ver: 2.0.10 #### Date: 24/01/2005  ########################
                rs_UsrAcc!UaValidCoCd = ""
                rs_UsrAcc!UaTag = Trim(wo_Ctl.Tag)
' ###########################################  Manoj  ###########################################
                rs_UsrAcc!UaTlBarDesc = ""  'Jen Clr Scr
                rs_UsrAcc!ModUsr = ""
                rs_UsrAcc!ModDt = MWLib.EmptyDate
                rs_UsrAcc!ModTime = 0
                rs_UsrAcc!UaSeoYN = ""
                rs_UsrAcc!UaReplNo = 0    '*** (Jen 2.14 Next Patch 1)
                'rs_UsrAcc!UaMetlrCd = ""   '***** Manali 3.8.0 - MetlrCd
                Debug.Print rs_UsrAcc!UaMnuCd + ", " + rs_UsrAcc!UaTag
                rs_UsrAcc.Update
            End If
        End If
    Next
    On Error GoTo IgnoreErrH
    '***** Manali 3.8.0 - UaTag added in case of isNull
    Dim wMsg As String
    wMsg = ""
    ADC.Connection.Execute ("Update UsrAccess Set UaTag = IsNull((Select isNull(UA1.UaTag , '') From UsrAccess UA1 " + _
        "Where UA1.UaMnuCd= UsrAccess.UaMnuCd And UA1.UaMnuForm = UsrAccess.UaMnuForm  And " + _
        "UA1.UAUSRCD= '" + ctGenMenuUsr + "'), UaTag) Where UAMNUFORM = '" + ms_MnuForm + "'")
IgnoreErrH:
    wMsg = "Please copy to individual Users"
    DispMsg "Menu on '" + ms_MnuForm + "' From Succesfully Generated for '" + ctGenMenuUsr + "'" + "." + wMsg, etInfo
    If Err.Number <> 0 Then
        MWLib.EndProcess Me
        Exit Sub
      End If
GenMnuErr:
    If Err.Number <> 0 Then
      DispMsg Err.Description + " Menu Can Not Generate", etError
      If Err.Number = -2147217900 Then rs_UsrAcc.CancelUpdate
    End If
    MWLib.EndProcess Me
End Sub
Private Sub HideAllMnus()
    Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption <> "DUM" Then
                Ctl.Visible = False
           End If
        End If
    Next
End Sub
Private Sub HideDumMnus()
    On Error GoTo ErrH
    Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption = "DUM" Then
                Ctl.Visible = False
           End If
        End If
    Next
    Exit Sub
ErrH:
    Resume Next
End Sub
Private Sub InitRtn()
    Set Ctl.ActScrn = Screen
    gs_CoCd = UCase(EmrFrmJMDI.TxtCoCd.text)
    gs_CoNm = ADC.Connection.GetFldVal("Select HName from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
    'gs_CoNm = ADC.Connection.GetFldVal("Select HName from Head where HCd='" + ctSelfCmCd + "'")

    If Dir(App.Path + "\Blue.Bmp") = "" Then
        MsgBox "Reqd Picture files missing.Contact SysAdmin"
        End
    End If
    StatusBar2.Panels("ErrPct").Picture = LoadPicture(App.Path + "\Blue.Bmp")
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If UnloadMode <> vbFormCode Then Cancel = True     'Original Code (Before Color)
    
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
End Sub
Private Function ChkInstance() As Boolean
    ChkInstance = True
    If App.PrevInstance Then
        MsgBox "Instance already running"
        ChkInstance = False
    End If
End Function
Private Sub StatusBar2_PanelClick(ByVal Panel As ComctlLib.Panel)
On Error GoTo PnlErrHndl
  'If Panel.Key = "ErrPct" Then MWLib.ViewError    'Original Code (Before Color)
  If Panel.Key = "ErrPct" Then MWLib.ViewError (IIF(Me.ActiveForm Is Nothing, Me, Me.ActiveForm))
  If Panel.Key = "InMail" And gb_MailAct = True Then Call SetMsgCnt(True)
  
  Exit Sub
PnlErrHndl:
  DispMsg Err.Description, etWarning
End Sub

' ########################  Manoj #### Ver: 2.0.6 #### Date: 16/06/2004  ########################
' Fetch the Host Machine Name
Private Function GetHostNm() As String
  Dim wsTmp As String
  wsTmp = String$(128, Chr(0))
  Call GetComputerName(wsTmp, 255)
  wsTmp = Replace(wsTmp, Chr(0), "")
  GetHostNm = wsTmp
End Function
' ###########################################  Manoj  ###########################################

' ########################  Manoj #### Ver: 2.11.0 #### Date: 27/05/2005  ########################
Private Sub MailLogin()
  Dim rsUM As MDORowSet, wsLoginNm As String, wsPass As String
  Set rsUM = ADC.Connection.OpenRes("Select * from UsrMail Where UmUsrCd = '" + gs_UsrCd + "'")
  If rsUM.RecCount > 0 Then
    wsLoginNm = rsUM.FldValue("UmMailLogin")
    wsPass = rsUM.FldValue("UmMailPass")
    If rsUM.FldValue("UmValidHostNm") = "" Or InStr(1, "," + rsUM.FldValue("UmValidHostNm") + ",", "," + GetHostNm + ",") > 0 Then
      If LogOnSes(wsLoginNm, wsPass) = True Then
        DispMsg "Mail Sign in Succeeded.", etInfo
        gb_MailAct = True
        Call SetMailMenu
      End If
    End If
  End If
End Sub
Private Function LogOnSes(ps_LoginNm As String, ps_Password As String) As Boolean
  If oMapiSes.NewSession Then
    MsgBox "Session already established"
    Exit Function
  End If
  On Error GoTo ErrLoginFail
  With oMapiSes
    ' Set DownLoadMail to False to prevent immediate download.
    .DownLoadMail = False
    .LogonUI = True ' Use the underlying email system's logon UI.
    .UserName = ps_LoginNm
    .Password = ps_Password
    .SignOn ' Signon method.
    ' If successful, return True
    LogOnSes = True
    ' Set NewSession to True and set0
    .NewSession = True
    ' variable flag to true
    oMapiMsg.SessionID = .SessionID ' You must set this before continuing.
    Call SetMsgCnt(True)
  End With

  Exit Function
ErrLoginFail:
  If Err.Number = 32003 Then
    DispMsg "Sign In Process Canceled", etWarning
  Else
    DispMsg Err.Description, etWarning
  End If
  Call LogOffSes
  LogOnSes = False
End Function
Private Function LogOffSes() As Boolean
    ' Logoff the MapSessions control.
    With oMapiSes
      If .NewSession = True Then
        .SignOff ' Close the session.
        .NewSession = False ' Flag for new session.
        gb_MailAct = False
        Call SetMailMenu
      End If
    End With
End Function
Private Function GetMsgCnt(ByVal pbUnRead As Boolean) As Integer
  oMapiMsg.FetchSorted = True
  If pbUnRead = True Then
    oMapiMsg.FetchUnreadOnly = True
  Else
    oMapiMsg.FetchUnreadOnly = False
  End If
  oMapiMsg.Fetch
  GetMsgCnt = oMapiMsg.MsgCount
  Exit Function
End Function
Public Sub SetMsgCnt(ByVal pbUnRead As Boolean)
  Dim wiMsgCnt  As Integer
  wiMsgCnt = GetMsgCnt(True)
  StatusBar2.Panels("InMail").text = IIF(wiMsgCnt <> 0, CStr(wiMsgCnt), "")
End Sub
Public Function ComposeMsg(Optional pb_NewMsg As Boolean = True, Optional pb_ShowWin As Boolean, _
                      Optional ps_RecpAddr As String, Optional ps_MsgSubj As String, Optional ps_MsgBody As String) As Boolean
  On Error GoTo ComposeErr
  Dim strMessage As String
  Dim ws_RecpAddr() As String, i As Integer
  ' Use the Compose method and then invoke the
  ' Send method. When the optional argument
  ' is set to True, the underlying mail system's
  ' form is used. Otherwise, you must create your
  ' own.
  If IsMissing(pb_ShowWin) = True Then pb_ShowWin = False
  If IsMissing(ps_RecpAddr) = True Then ps_RecpAddr = ""
  If IsMissing(ps_MsgSubj) = True Then ps_MsgSubj = ""
  If IsMissing(ps_MsgBody) = True Then ps_MsgBody = ""

  With oMapiMsg
    .Compose
    If pb_NewMsg = False Then
      ws_RecpAddr = Split(ps_RecpAddr, ",")
      For i = LBound(ws_RecpAddr) To UBound(ws_RecpAddr)
        .RecipIndex = i
        .RecipAddress = ws_RecpAddr(i)
        .AddressResolveUI = True
        .ResolveName
        .RecipType = mapToList
      Next
      .MsgSubject = ps_MsgSubj
      .MsgNoteText = ps_MsgBody
    End If
    oMapiMsg.send IIF(pb_NewMsg Or pb_ShowWin, True, False)
  End With
  ComposeMsg = True
  Exit Function
ComposeErr:
  If Err.Number = 32003 Then
    DispMsg "Login Terminated Abnormally.", etWarning
    Call LogOffSes
  Else
    DispMsg Err.Description, etWarning
  End If
  ComposeMsg = False
End Function

Private Sub MDIForm_Activate()
On Error GoTo MsgCntHndl
  If gb_MailAct = True Then
    Call SetMsgCnt(True)
  End If
  Exit Sub
MsgCntHndl:
  DispMsg Err.Description, etWarning
End Sub
Private Sub SetMailMenu()
  If MailSignIn.Visible = True Then MailSignIn = Not gb_MailAct
  If MailSignOut.Visible = True Then MailSignOut = gb_MailAct
  If SndMail.Visible = True Then SndMail = gb_MailAct
  StatusBar2.Panels("InMail").Visible = gb_MailAct
End Sub
' ###########################################  Manoj  ###########################################

' Procedure To Get Tag & Code of Menu
Private Sub PrnMnus()
    Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption <> "DUM" Then
                Debug.Print Ctl.Tag + vbTab + vbTab + vbTab + Ctl.Name
           End If
        End If
    Next
End Sub

'*** Jenny Color
Private Sub MnuTlBar_Click()
  If TlBar.Visible = True Then
    TlBar.Visible = False
    'MnuTlBar.Caption = "&ToolBar"
  Else
    TlBar.Top = 0
    If CmdCtlToolBar.Count = 1 And CmdCtlToolBar(0).Tag = "" Then Exit Sub
    TlBar.Visible = True
    'MnuTlBar.Caption = "&Close ToolBar"
  End If
End Sub
Private Sub CreateToolBar()
  '*** (Bef 21/05/05) Set mo_CsTlBarPic = GetRecClrSet.FldPicValue("CsTlBarPic")
  
  '*** (Jen 2.13)
  If Not (GetRecClrSet Is Nothing) Then
  '*** (Jen 2.13)
    '*** (21/05/05)
    'If Trim(GetRecClrSet!CsTlBarPic) <> "" And Dir(Trim(GetRecClrSet!CsTlBarPic)) <> "" Then
    If Trim(GetRecClrSet!CsTlBarPic) <> "" And IsValidPath(GetClrPth & Trim(GetRecClrSet!CsTlBarPic)) = True Then
      Set mo_CsTlBarPic = LoadPicture(GetClrPth & GetRecClrSet!CsTlBarPic)
    Else
      Set mo_CsTlBarPic = CmdCtlToolBar(0).Pic
    End If
    '*** (21/05/05)
  '*** (Jen 2.13)
  Else
    Set mo_CsTlBarPic = LoadPicture("")
  End If
  '*** (Jen 2.13)
    
  Set mo_TlBarColl = New Collection
  
  Dim btnX As Button
  Set btnX = TlBar.Buttons.Add(, , , tbrSeparator)
  
  'CreateToolBarBtn "SO", "SalOrd", 0, "Sales Order"
  'CreateToolBarBtn "WIW", "RepWIWPrd", 1, "What Is Where"
  'CreateToolBarBtn "WIP", "RepWipBPcs", 2, "Work In Progress (Bag)"
  'Set mo_CsTlBarPic = Nothing
End Sub
Private Sub CreateToolBarBtn(ps_Caption As String, ps_MnuCd As String, pi_indx As Integer, ps_Desc As String)
  Dim btnX As Button
  
  Set btnX = TlBar.Buttons.Add(, ps_MnuCd, , tbrPlaceholder)
  '*** (Bef 03/08/05) btnX.Width = GetRecClrSet!CsTlBarBtnWidth + GetRecClrSet!CsTlBarBtnGap '*** Placeholder width to accommodate a CommandButton.
  '*** (Jen 2.13)
  If Not (GetRecClrSet Is Nothing) Then
  '*** (Jen 2.13)
    '*** (03/08/05)
    If GetRecClrSet!CsTlBarBtnWidth > 0 Then
      btnX.Width = GetRecClrSet!CsTlBarBtnWidth + GetRecClrSet!CsTlBarBtnGap '*** Placeholder width to accommodate a CommandButton.
    Else
      btnX.Width = CmdCtlToolBar(0).Width + GetRecClrSet!CsTlBarBtnGap
    End If
    '*** (03/08/05)
  
  '*** (Jen 2.13)
  Else
    btnX.Width = CmdCtlToolBar(0).Width
  End If
  '*** (Jen 2.13)
  
  If pi_indx > 0 Then
    '*** For creating the next element in the control array
    Load CmdCtlToolBar(pi_indx)
  End If
  
  '*** Configure CommandButton control to be at same location as the
  '    Button object with the PlaceHolder style (key = ps_Caption).
  With CmdCtlToolBar(pi_indx)
    .Pic = mo_CsTlBarPic
    '*** (Jen 2.13)
    If Not (GetRecClrSet Is Nothing) Then
    '*** (Jen 2.13)
      .ForeColor = Val(GetRecClrSet!CsTlBarBtnForClr)   'vbBlack
      .DisableColor = Val(GetRecClrSet!CsTlBarBtnDisaClr)
      .MaskColor = Val(GetRecClrSet!CsTlBarBtnMaskClr)
    '*** (Jen 2.13)
    End If
    '*** (Jen 2.13)
    .Caption = ps_Caption
    .Visible = True
    .Width = TlBar.Buttons(ps_MnuCd).Width
    '*** (Bef 03/08/05)
    '.Height = GetRecClrSet!CsTlBarBtnHeight
    '*** (Bef 03/08/05)
    
    '*** (Jen 2.13)
    If Not (GetRecClrSet Is Nothing) Then
    '*** (Jen 2.13)
      '*** 03/08/05
      If GetRecClrSet!CsTlBarBtnHeight > 0 Then
        .Height = GetRecClrSet!CsTlBarBtnHeight
      Else
        .Height = CmdCtlToolBar(0).Height
      End If
      If GetRecClrSet!CsTlBarBtnWidth <= 0 Or GetRecClrSet!CsTlBarBtnHeight <= 0 Then
        .ReSizePic = True
      End If
      '*** 03/08/05
    
    '*** (Jen 2.13)
    Else
      .Height = CmdCtlToolBar(0).Height
    End If
    '*** (Jen 2.13)
    
    .Top = TlBar.Buttons(ps_MnuCd).Top
    .Left = TlBar.Buttons(ps_MnuCd).Left
   
    .ToolTipTxt = ps_Desc
    .Tag = ps_MnuCd
    mo_TlBarColl.Add pi_indx, UCase(ps_MnuCd)
  End With
End Sub
Private Sub CmdCtlToolBar_Click(Index As Integer)
  FormSelect CmdCtlToolBar(Index).Tag
  'ShowForm ?, CmdCtlToolBar(Index).Tag
End Sub
Private Sub SetMDIPic()
  '*** The Component 'Microsoft PictureClip Control 6.0 (SP3)'  (File Name: PICCLP32.OCX) has been added to the project

  '*** MSDN Help Topics
  ' PictureClip Control Example
  ' HOWTO: Create a Resolution-Independent Form
  '*** MSDN Help Topics

  If mo_MDIPic Is Nothing Then Exit Sub
  Dim xTwips, yTwips, xPixels, yPixels

  xTwips = Screen.TwipsPerPixelX
  yTwips = Screen.TwipsPerPixelY
  xPixels = Screen.Width / xTwips  '*** X Pixel Resolution
  yPixels = Screen.Height / yTwips '*** Y Pixel Resolution

  Set PictureClip1.Picture = mo_MDIPic
  If PictureClip1.StretchX <> xPixels Or PictureClip1.StretchY <> yPixels Then

    '*** Get X and Y coordinates of the clipping region.
    PictureClip1.ClipX = 0
    PictureClip1.ClipY = 0

    '*** Set the area of the clipping region (in pixels).
    PictureClip1.ClipHeight = PictureClip1.Height
    PictureClip1.ClipWidth = PictureClip1.Width

    '*** Set the destination area to fill the picture box.
    If PictureClip1.StretchX <> xPixels Then PictureClip1.StretchX = xPixels
    If PictureClip1.StretchY <> yPixels Then PictureClip1.StretchY = yPixels

    '*** Assign the clipped bitmap to the picture box.
    Me.Picture = PictureClip1.Clip

  End If
End Sub
'*** Jenny Color
'Private Sub FormSelect(ByVal ps_Mnu As String)
'  '*** Jenny Color
'
'  gs_MnuCd = ps_Mnu   '*** 13/06/05
'
'  'Dim wfrm As New EmrFrmCustMst
'  Dim wfrm As Form
'  Set wfrm = Nothing
'
'  If GetRecClrSet.RecCount > 0 Then
'    GetRecClrSet.MoveFirst
'    '*** For Setting System Colors (For Tab Page)
'    SetTabSysColors Val(GetRecClrSet!CsTabPgClr), Val(GetRecClrSet!CsTabForClr), Val(GetRecClrSet!CsTabBorderClr), Val(GetRecClrSet!CsTabShadowClr), Val(GetRecClrSet!CsTabHighLightClr)
'    '*** For Setting System Colors (For Tab Page)
'  End If
'
'  Select Case UCase(ps_Mnu)
'  Case Is = UCase("Head")
'    Set wfrm = New EmrFrmHead
'  Case Is = UCase("CustMst")
'    Set wfrm = New EmrFrmCustMst
'  Case Is = UCase("Param")
'    Set wfrm = New EmrFrmParam
'    wfrm.Caption = "Parameters"
'  Case Is = UCase("SalOrd")
'    Set wfrm = New EmrFrmOrdMst
'  Case Is = UCase("DM")
'    Set wfrm = New EmrFrmDsgMst
'  Case Is = UCase("PM")
'    Set wfrm = New EmrFrmDsgMst
'  Case Is = UCase("SM")
'    Set wfrm = New EmrFrmDsgMst
'  Case Is = UCase("VarParam")
'    Set wfrm = New EmrFrmvParam
'    wfrm.Caption = "Variable Parameters"
'  Case Is = UCase("Bag")
'    Set wfrm = New EmrFrmBag
'  Case Is = UCase("RC")
'    Set wfrm = New EmrFrmBagReco
'    wfrm.Caption = "Bag Reconciliation"
'  Case Is = UCase("Bnk")
'    Set wfrm = New EmrFrmBnkAg
'  Case Is = UCase("CpyRt")
'    Set wfrm = New EmrFrmCpyRt
'  Case Is = UCase("InCustmMtch")
'    Set wfrm = New EmrFrmCustmMatch
'    wfrm.Caption = "Invoice Custom Matching"
'  Case Is = UCase("Tz")
'    Set wfrm = New EmrFrmCustmPur
'    wfrm.Caption = "Custom Purchase"
'  Case Is = UCase("CustMst")
'    Set wfrm = New EmrFrmCustMst
'  Case Is = UCase("SuppMst")
'    Set wfrm = New EmrFrmCustMst
'  Case Is = UCase("PLMst")
'    Set wfrm = New EmrFrmCustMst
'  Case Is = UCase("CoMst")
'    Set wfrm = New EmrFrmCustMst
'    wfrm.Caption = "CoMst"
'  Case Is = UCase("ContractorMst")
'    Set wfrm = New EmrFrmCustMst
'  Case Is = UCase("DtTbl")
'    Set wfrm = New EmrFrmDtTbl
'  Case Is = UCase("FB")
'    Set wfrm = New EmrFrmFg
'    wfrm.Caption = "Finished Goods (Bag)"
'  Case Is = UCase("FM")
'    Set wfrm = New EmrFrmFgMvm
'    wfrm.Caption = "Finished Goods (Bag Movement)"
'  Case Is = UCase("FWB")
'    Set wfrm = New EmrFrmFg
'    wfrm.Caption = "Finished Goods (Bag Return)"
'  Case Is = UCase("Head")
'    Set wfrm = New EmrFrmHead
'  Case Is = UCase("InAllBags")
'    Set wfrm = New EmrFrmInvAllBags
'    wfrm.Caption = "Invoice All Bags"
'  Case Is = UCase("IN")
'    Set wfrm = New EmrFrmInvHd
'    wfrm.Caption = "Invoice Header"
'  Case Is = UCase("LabRt")
'    Set wfrm = New EmrFrmLabRt
'  Case Is = UCase("RepLst")
''    *** Form with Menu Placed for all Listing Reports
'    Set wfrm = New EmrFrmList
'  Case Is = UCase("Loc")
'    Set wfrm = New EmrFrmLoc
'  Case Is = UCase("LossEntry")
'    Set wfrm = New EmrFrmLossEntry
'  Case Is = UCase("LsRec")
'    Set wfrm = New EmrFrmLsRec
'  Case Is = UCase("Zoom")
'    Set wfrm = New EmrFrmOrdZoom
'  Case Is = UCase("PpcCap")
'    Set wfrm = New EmrFrmPpcCap
'  Case Is = UCase("PpcDsgLd")
'    Set wfrm = New EmrFrmPpcDsgLd
'  Case Is = UCase("PpcOrdPln")
'    Set wfrm = New EmrFrmPpcOrdPln
'  Case Is = UCase("PrdPts")
'    Set wfrm = New EmrFrmPrdPts
'  Case Is = UCase("ResAuto")
'    Set wfrm = New EmrFrmResAutogen
'  Case Is = UCase("RmMst")
'    Set wfrm = New EmrFrmRmMst
'    wfrm.Caption = "Raw Material Master"
'  Case Is = UCase("RmRt")
'    Set wfrm = New EmrFrmRmRt
'  Case Is = UCase("RmRtHist")
'    Set wfrm = New EmrFrmRmRt
'  Case Is = UCase("RepTagPrn")
'    Set wfrm = New EmrFrmTag
'    wfrm.Caption = "Tag Printing"
'  Case Is = UCase("TagFmt")
'    If adc.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") _
'      And _
'      adc.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmtSr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") Then
'        '  If gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0" Then
'      Set wfrm = New EmrFrmTagFmt
'    End If
'  Case Is = UCase("Tree")
'    Set wfrm = New EmrFrmTree
'  Case Is = UCase("XMLToEmr")
'    Set wfrm = New EmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Order)"
'  Case Is = UCase("EmrToXML")
'    Set wfrm = New EmrFrmTrfOut
'  Case Is = UCase("XMLMstToEmr")
'    Set wfrm = New xEmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Master)"
'  Case Is = UCase("SA")
'    Set wfrm = New EmrFrmTxnRm
'    wfrm.Caption = "Sales"
'  Case Is = UCase("CNV")
'    Set wfrm = New EmrFrmTxnRm
'    wfrm.Caption = "Conversion"
'  Case Is = UCase("PR")
'    Set wfrm = New EmrFrmTxnRm
'    wfrm.Caption = "Purchase"
'  Case Is = UCase("MV")
'    Set wfrm = New EmrFrmTxnRm
'    wfrm.Caption = "Raw Material Issue"
'  Case Is = UCase("UsrAcss")
'    Set wfrm = New EmrFrmUsrAcss
'    wfrm.Caption = "User Access Master"
'  Case Is = UCase("UsrAdm")
'    Set wfrm = New EmrFrmUsrAdm
'    wfrm.Caption = "User Administration"
'  Case Is = UCase("UsrMail")
'    Set wfrm = New EmrFrmUsrMail
'    wfrm.Caption = "User Mail"
'  Case Is = UCase("VarParam")
'    Set wfrm = New EmrFrmvParam
'    wfrm.Caption = "Variable Parameters"
'  Case Is = UCase("XMLMstToEmr")
'    Set wfrm = New xEmrFrmTrfIn
'    wfrm.Caption = "XML To Emr (Master)"
'  Case Is = UCase("EmrUpd")
'    Set wfrm = New EmrSuppFrmUpd
'  Case Is = UCase("RepSOPnd")
'    Set wfrm = New EmrFrpSOPnd
'    wfrm.Caption = "Sales Order Pending"
'  Case Is = UCase("RepBagHist")
'    Set wfrm = New EmrFrpBagHist
'    wfrm.Caption = "Bag History"
'  Case Is = UCase("RmZMst")
'    Set wfrm = New EmrFrmRmMst
'    wfrm.Caption = "Customs Raw Material Master"
'  Case Is = UCase("PrcLst")
'    Set wfrm = New EmrFrmOrdMst
'  Case Is = UCase("Quot")
'    Set wfrm = New EmrFrmOrdMst
'  Case Is = UCase("QB")
'    Set wfrm = New EmrFrmOrdMst
'  Case Is = UCase("SkQuot")
'    Set wfrm = New EmrFrmOrdMst
'  Case Is = UCase("BM")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Bag Merge"
'  Case Is = UCase("DT")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Daily Transactions"
'  Case Is = UCase("BV")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Bag Movement"
'  Case Is = UCase("BA")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Bag Acknowledgement"
'  Case Is = UCase("BS")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Bag Split"
'  Case Is = UCase("REJ")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Rejection"
'  Case Is = UCase("MLT")
'    Set wfrm = New EmrFrmTxn
'    wfrm.Caption = "Melting"
'  Case Is = UCase("CTBIn")
'    Set wfrm = New EmrFrmInvHd
'    wfrm.Caption = "CTB Header"
'  Case Is = UCase("CTBInAllBags")
'    Set wfrm = New EmrFrmInvAllBags
'    wfrm.Caption = "CTB All Bags"
'  Case Is = UCase("CTBInDsg")
'    Set wfrm = New EmrFrmInvDsg
'    wfrm.Caption = "CTB Design"
'  Case Is = UCase("RepBagPrn")
'    Set wfrm = New EmrFrpBagPrn
'    wfrm.Caption = "Bag Printing"
'  Case Is = UCase("RepBagReco")
'    Set wfrm = New EmrFrpBagReco
'    wfrm.Caption = "Bag Reconciliation Report"
'  Case Is = UCase("RepBrkMsgRm")
'    Set wfrm = New EmrFrpBrkMsg
'    wfrm.Caption = "Breakage, Missing Report (Base= Rm)"
'  Case Is = UCase("RepBrkMsgPrd")
'    Set wfrm = New EmrFrpBrkMsg
'    wfrm.Caption = "Breakage, Missing Report (Base= Prd)"
'  Case Is = UCase("RepBv")
'    Set wfrm = New EmrFrpBv
'    wfrm.Caption = "Bag Movement (Bag Pieces)"
'  Case Is = UCase("RepBvRm")
'    Set wfrm = New EmrFrpBvRm
'    wfrm.Caption = "Bag Movement (Raw Material)"
'  Case Is = UCase("RepCustmStk")
'    Set wfrm = New EmrFrpCustmStk
'    wfrm.Caption = "Custom Stock"
'  Case Is = UCase("RepDiaProfit")
'    Set wfrm = New EmrFrpDiaProfit
'    wfrm.Caption = "Diamond Profit"
'  Case Is = UCase("InDsg")
'    Set wfrm = New EmrFrmInvDsg
'    wfrm.Caption = "Invoice Design"
'  Case Is = UCase("RepFlashStk")
'    Set wfrm = New EmrFrpFlashStk
'    wfrm.Caption = "Flash Stock"
'  Case Is = UCase("RepFltBagPrn")
'    Set wfrm = New EmrFrpFltBagPrn
'    wfrm.Caption = "Flute Bag Printing"
'  Case Is = UCase("RepGldLs")
'    Set wfrm = New EmrFrpGldLs
'    wfrm.Caption = "Gold Loss"
'  Case Is = UCase("RepInv")
'    Set wfrm = New EmrFrpInv
'    wfrm.Caption = "Invoice Printing"
'  Case Is = UCase("RepGrFormFrnt")
'    Set wfrm = New EmrFrpInv
'    wfrm.Caption = "GR Form (Front)"
'  Case Is = UCase("RepShpBill")
'    Set wfrm = New EmrFrpInv
'    wfrm.Caption = "Shipping Bill"
'  Case Is = UCase("RepGSPFORM")
'    Set wfrm = New EmrFrpInv
'    wfrm.Caption = "GSP Form"
'  Case Is = UCase("RepInvBOELst")
'    Set wfrm = New EmrFrpInvBoeLst
'    wfrm.Caption = "Invoice BOE Listing"
'  Case Is = UCase("RepDetInv")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "Invoice Costing"
'  Case Is = UCase("RepAvgInv")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "Average Invoice Printing"
'  Case Is = UCase("RepInvCat")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "Invoice Catalogue"
'  Case Is = UCase("RepDetInvCTB")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "CTB Invoice Costing"
'  Case Is = UCase("RepAvgInvCTB")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "Average CTB Invoice Printing"
'  Case Is = UCase("RepInvCTBCat")
'    Set wfrm = New EmrFrpInvCost
'    wfrm.Caption = "CTB Invoice Catalogue"
'  Case Is = UCase("RepInvCustStmt")
'    Set wfrm = New EmrFrpInvCustStmt
'  Case Is = UCase("RepGrFormBack")
'    Set wfrm = New EmrFrpInvLet
'    wfrm.Caption = "GR Form (Back)"
'  Case Is = UCase("RepGSPDetLet")
'    Set wfrm = New EmrFrpInvLet
'    wfrm.Caption = "GSP Detail Letter"
'  Case Is = UCase("RepGSPCert")
'    Set wfrm = New EmrFrpInvLet
'    wfrm.Caption = "GSP Certification"
'  Case Is = UCase("RepDeclnLet")
'    Set wfrm = New EmrFrpInvLet
'    wfrm.Caption = "Declaration Letter"
'  Case Is = UCase("RepInvLossSubRep")
'    Set wfrm = New EmrFrpInvLossSubRep
'    wfrm.Caption = "Invoice Metal Loss Sub Report"
'  Case Is = UCase("RepInvProfitMrg")
'    Set wfrm = New EmrFrpInvProfitMrg
'    wfrm.Caption = "Profit Margin"
'  Case Is = UCase("RepLsRec")
'    Set wfrm = New EmrFrpLsRec
'  Case Is = UCase("RepNetStk")
'    Set wfrm = New EmrFrpNetStk
'  Case Is = UCase("RepDetOrd")
'    Set wfrm = New EmrFrpOrd
'    wfrm.Caption = "Detail Order Printing"
'  Case Is = UCase("RepAvgOrd")
'    Set wfrm = New EmrFrpOrd
'    wfrm.Caption = "Average Order Printing"
'  Case Is = UCase("RepOrdCat")
'    Set wfrm = New EmrFrpOrd
'    wfrm.Caption = "Order Catalogue Printing"
'  Case Is = UCase("RepOrdProfitMrg")
'    Set wfrm = New EmrFrpOrdProfitMrg
'    wfrm.Caption = "Order Profit Margin"
'  Case Is = UCase("RepPckLst")
'    Set wfrm = New EmrFrpPckLst
'    wfrm.Caption = "Finished Goods Packing List Fr INV"
'  Case Is = UCase("RepPckLstInv")
'    Set wfrm = New EmrFrpPckLstInv
'    wfrm.Caption = "Invoice Packing List"
'  Case Is = UCase("RepCntrlCpy")
'    Set wfrm = New EmrFrpPpcCntrlCpy
'  Case Is = UCase("RepOrdPln")
'    Set wfrm = New EmrFrpPpcOrdPln
'  Case Is = UCase("RepDsgLd")
'    Set wfrm = New EmrFrpPpcDsgLd
'  Case Is = UCase("RepPrd")
'    Set wfrm = New EmrFrpPrd
'    wfrm.Caption = "Production Report"
'  Case Is = UCase("RepPrdPts")
'    Set wfrm = New EmrFrpPrdPts
'    wfrm.Caption = "Production Points Report"
'  Case Is = UCase("RepRejAna")
'    Set wfrm = New EmrFrpRejAna
'    wfrm.Caption = "Rejection Analysis"
'  Case Is = UCase("RepRmMst")
'    Set wfrm = New EmrFrpRmMst
'    wfrm.Caption = "Raw Material/ Custom Master Listing"
'  Case Is = UCase("RepRmReqStat")
'    Set wfrm = New EmrFrpRmReq
'    wfrm.Caption = "Raw Material Requirement (Static)"
'  Case Is = UCase("RepRmReqDyn")
'    Set wfrm = New EmrFrpRmReq
'    wfrm.Caption = "Raw Material Requirement (Dynamic)"
'  Case Is = UCase("RepRmReqLoc")
'    Set wfrm = New EmrFrpRmReqLoc
'    wfrm.Caption = "Raw Material Requirement For Loc"
'  Case Is = UCase("RepRmRt")
'    Set wfrm = New EmrFrpRmRt
'    wfrm.Caption = "Raw Material Rate Master Listing"
'  Case Is = UCase("RepLabRt")
'    Set wfrm = New EmrFrpLabRt
'    wfrm.Caption = "Labour Rate Master Listing"
'  Case Is = UCase("RepRmRtHist")
'    Set wfrm = New EmrFrpRmRt
'    wfrm.Caption = "Raw Material Rate History Listing"
'  Case Is = UCase("RepLoc")
'    Set wfrm = New EmrFrpLoc
'  Case Is = UCase("RepSalAnaDsg")
'    Set wfrm = New EmrFrpSalAnaDsg
'    wfrm.Caption = "Sales Analysis (Design Level)"
'  Case Is = UCase("RepSalAnaRm")
'    Set wfrm = New EmrFrpSalAnaRm
'    wfrm.Caption = "Sales Analysis (Rm Level)"
'  Case Is = UCase("RepSalAnaLab")
'    Set wfrm = New EmrFrpSalAnaLab
'    wfrm.Caption = "Sales Analysis (Lab Level)"
'  Case Is = UCase("RepStkLed")
'    Set wfrm = New EmrFrpStkLed
'    wfrm.Caption = "Stock Ledger"
'  Case Is = UCase("RepStkSumm")
'    Set wfrm = New EmrFrpStkSumm
'    wfrm.Caption = "Stock Summary"
'  Case Is = UCase("RepAnnexValAddn")
'    Set wfrm = New EmrFrpValAddnAnnex
'    wfrm.Caption = "Annexure To Value Addition"
'  Case Is = UCase("RepValAddn")
'    Set wfrm = New EmrFrpValAddn
'    wfrm.Caption = "Value Addition"
'  Case Is = UCase("RepWipBPcs")
'    Set wfrm = New EmrFrpWipBPcs
'    wfrm.Caption = "WIP Bag Pieces Report"
'  Case Is = UCase("RepWipFgBPcs")
'    Set wfrm = New EmrFrpWipBPcs
'    wfrm.Caption = "Fg Bag Pieces Report"
'  Case Is = UCase("RepWipBPcsBkDt")
'    Set wfrm = New EmrFrpWipBPcsBkDt
'    wfrm.Caption = "WIP Bag Pieces Report (Back Dated)"
'  Case Is = UCase("RepWipFgBPcsBkDt")
'    Set wfrm = New EmrFrpWipBPcsBkDt
'    wfrm.Caption = "Fg Bag Pieces Report (Back Dated)"
'  Case Is = UCase("RepWipBRm")
'    Set wfrm = New EmrFrpWipBRm
'    wfrm.Caption = "WIP Bag Raw Material Report"
'  Case Is = UCase("RepWipFgBRm")
'    Set wfrm = New EmrFrpWipBRm
'    wfrm.Caption = "Fg Bag Raw Material Report"
'  Case Is = UCase("RepWIWPrd")
'    Set wfrm = New EmrFrpWIW
'    wfrm.Caption = "What Is Where (Production)"
'  Case Is = UCase("RepWIWExp")
'    Set wfrm = New EmrFrpWIW
'    wfrm.Caption = "What Is Where (Export)"
'  Case Is = UCase("RepWIWLongPrd")
'    Set wfrm = New EmrFrpWIW
'    wfrm.Caption = "What Is Where Long (Production)"
'  Case Is = UCase("RepWIWLongExp")
'    Set wfrm = New EmrFrpWIW
'    wfrm.Caption = "What Is Where Long (Export)"
'
'
'
'
'  'Case Is = UCase("ClrDfn")
'  '  Set wfrm = New EmrFrmClrDfn
'  'Case Is = UCase("ClrSel")
'  '  Set wfrm = New EmrFrmClrPic
'
'
'
'
'  '##################################################################################
'  '##################################################################################
''  Case Is = UCase("BM")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Bag Merge"
''  Case Is = UCase("EmrToXML")
''    Set wfrm = New EmrFrmTrfOut
''  Case Is = UCase("EmrUpd")
''    Set wfrm = New EmrSuppFrmUpd
''  Case Is = UCase("FM")
''    Set wfrm = New EmrFrmFgMvm
''    wfrm.Caption = "Finished Goods (Bag Movement)"
''    'wfrm.WindowState = 2
''  Case Is = UCase("Head")
''    Set wfrm = New EmrFrmHead
''    'wfrm.WindowState = 2
''  Case Is = UCase("CustMst")
''    'Dim wfrm As New EmrFrmChange
''    'Dim wfrm As New EmrFrmPpcCap
''    'Dim wfrm As New EmrFrmPpcDsgLd
''    Set wfrm = New EmrFrmCustMst
''  Case Is = UCase("RC")
''    Set wfrm = New EmrFrmBagReco
''    wfrm.Caption = "Bag Reconciliation"
''  Case Is = UCase("RepBagReco")
''    Set wfrm = New EmrFrpBagReco
''    wfrm.Caption = "Bag Reconciliation Report"
''  Case Is = UCase("RepBvRm")
''    Set wfrm = New EmrFrpBvRm
''    wfrm.Caption = "Bag Movement (Raw Material)"
''  Case Is = UCase("RepFltBagPrn")
''    Set wfrm = New EmrFrpFltBagPrn
''    wfrm.Caption = "Flute Bag Printing"
''  Case Is = UCase("RepLst")
''    '*** Form with Menu Placed for all Listing Reports
''    Set wfrm = New EmrFrmList
''  Case Is = UCase("RepRmStkReq")
''    Set wfrm = New EmrFrpRmStkReq
''    wfrm.Caption = "Raw Material Stock Requirement"
''  Case Is = UCase("RepNetStk")
''    Set wfrm = New EmrFrpNetStk
''  Case Is = UCase("RmRtHist")
''    Set wfrm = New EmrFrmRmRt
''  Case Is = UCase("SA")
''    Set wfrm = New EmrFrmTxnRm
''    wfrm.Caption = "Sales"
''  Case Is = UCase("SuppMst")
''    Set wfrm = New EmrFrmCustMst
''  Case Is = UCase("PLMst")
''    Set wfrm = New EmrFrmCustMst
''  Case Is = UCase("CoMst")
''    Set wfrm = New EmrFrmCustMst
''    wfrm.Caption = "CoMst"
''  Case Is = UCase("ContractorMst")
''    Set wfrm = New EmrFrmCustMst
''  Case Is = UCase("Param")
''    Set wfrm = New EmrFrmParam
''    wfrm.Caption = "Parameters"
''  Case Is = UCase("TagFmt")
''    If adc.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmt]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") _
''      And _
''      adc.Connection.RecSeek("select 1 from dbo.sysobjects where id = object_id(N'[TagFmtSr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1") Then
''        '  If gs_Ver = "2.6.0" Or gs_Ver = "2.7.0" Or gs_Ver = "2.8.0" Or gs_Ver = "2.9.0" Then
''      Set wfrm = New EmrFrmTagFmt
''    End If
''  Case Is = UCase("Tree")
''    Set wfrm = New EmrFrmTree
''  Case Is = UCase("VarParam")
''    Set wfrm = New EmrFrmvParam
''    wfrm.Caption = "Variable Parameters"
''  Case Is = UCase("UsrAcss")
''    Set wfrm = New EmrFrmUsrAcss
''    wfrm.Caption = "User Access Master"
''  Case Is = UCase("UsrAdm")
''    Set wfrm = New EmrFrmUsrAdm
''    wfrm.Caption = "User Administration"
''  Case Is = UCase("UsrMail")
''    Set wfrm = New EmrFrmUsrMail
''    wfrm.Caption = "User Mail"
''  Case Is = UCase("RmMst")
''    Set wfrm = New EmrFrmRmMst
''    wfrm.Caption = "Raw Material Master"
''  Case Is = UCase("RmZMst")
''    Set wfrm = New EmrFrmRmMst
''    wfrm.Caption = "Customs Raw Material Master"
''  Case Is = UCase("DM")
''    Set wfrm = New EmrFrmDsgMst
''  Case Is = UCase("PM")
''    Set wfrm = New EmrFrmDsgMst
''  Case Is = UCase("SM")
''    Set wfrm = New EmrFrmDsgMst
''  Case Is = UCase("RmRt")
''    Set wfrm = New EmrFrmRmRt
''  Case Is = UCase("LabRt")
''    Set wfrm = New EmrFrmLabRt
''  Case Is = UCase("Loc")
''    Set wfrm = New EmrFrmLoc
''  Case Is = UCase("DtTbl")
''    Set wfrm = New EmrFrmDtTbl
''  Case Is = UCase("Bnk")
''    'Dim wfrm As New EmrFrmBnk
''    'Dim wfrm As New EmrFrmBnkOld
''    Set wfrm = New EmrFrmBnkAg
''  Case Is = UCase("Bag")
''    Set wfrm = New EmrFrmBag
''  Case Is = UCase("SalOrd")
''    Set wfrm = New EmrFrmOrdMst
''  Case Is = UCase("PrcLst")
''    Set wfrm = New EmrFrmOrdMst
''  Case Is = UCase("Quot")
''    Set wfrm = New EmrFrmOrdMst
''  Case Is = UCase("QB")
''    Set wfrm = New EmrFrmOrdMst
''  Case Is = UCase("SkQuot")
''    Set wfrm = New EmrFrmOrdMst
''  Case Is = UCase("ResAuto")
''    Set wfrm = New EmrFrmResAutogen
''  Case Is = UCase("CpyRt")
''    Set wfrm = New EmrFrmCpyRt
''  Case Is = UCase("DT")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Daily Transactions"
''    'wfrm.WindowState = 2
''  Case Is = UCase("BV")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Bag Movement"
''    'wfrm.WindowState = 2
''  Case Is = UCase("BA")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Bag Acknowledgement"
''    'wfrm.WindowState = 2
''  Case Is = UCase("BS")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Bag Split"
''    'wfrm.WindowState = 2
''  Case Is = UCase("REJ")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Rejection"
''    'wfrm.WindowState = 2
''  Case Is = UCase("MLT")
''    Set wfrm = New EmrFrmTxn
''    wfrm.Caption = "Melting"
''    'wfrm.WindowState = 2
''  Case Is = UCase("PrdPts")
''    Set wfrm = New EmrFrmPrdPts
''  Case Is = UCase("LossEntry")
''    Set wfrm = New EmrFrmLossEntry
''  Case Is = UCase("PR")
''    Set wfrm = New EmrFrmTxnRm
''    wfrm.Caption = "Purchase"
''    'wfrm.WindowState = 2
''  Case Is = UCase("Tz")
''    Set wfrm = New EmrFrmCustmPur
''    wfrm.Caption = "Custom Purchase"
''  Case Is = UCase("MV")
''    Set wfrm = New EmrFrmTxnRm
''    wfrm.Caption = "Raw Material Issue"
''    'wfrm.WindowState = 2
''  Case Is = UCase("CNV")
''    Set wfrm = New EmrFrmTxnRm
''    wfrm.Caption = "Conversion"
''    'wfrm.WindowState = 2
''  Case Is = UCase("FB")
''    Set wfrm = New EmrFrmFg
''    wfrm.Caption = "Finished Goods (Bag)"
''    'wfrm.WindowState = 2
''  Case Is = UCase("FWB")
''    Set wfrm = New EmrFrmFg
''    wfrm.Caption = "Finished Goods (Bag Return)"
''    'wfrm.WindowState = 2
''  Case Is = UCase("RepInvCustStmt")
''    Set wfrm = New EmrFrpInvCustStmt
''  Case Is = UCase("RepTreePerform")
''    Set wfrm = New EmrFrpTreeLst
''    wfrm.Caption = "Tree Performance Report"
''    wfrm.WindowState = 2
''  Case Is = UCase("RepTreeInvst")
''    Set wfrm = New EmrFrpTreeLst
''    wfrm.Caption = "Tree Investment Report"
''    wfrm.WindowState = 2
''  Case Is = UCase("IN")
''    Set wfrm = New EmrFrmInvHd
''    wfrm.Caption = "Invoice Header"
''  Case Is = UCase("InAllBags")
''    Set wfrm = New EmrFrmInvAllBags
''    wfrm.Caption = "Invoice All Bags"
''  Case Is = UCase("InDsg")
''    Set wfrm = New EmrFrmInvDsg
''    wfrm.Caption = "Invoice Design"
''  Case Is = UCase("InCustmMtch")
''    Set wfrm = New EmrFrmCustmMatch
''    wfrm.Caption = "Invoice Custom Matching"
''  Case Is = UCase("CTBIn")
''    Set wfrm = New EmrFrmInvHd
''    wfrm.Caption = "CTB Header"
''  Case Is = UCase("CTBInAllBags")
''    Set wfrm = New EmrFrmInvAllBags
''    wfrm.Caption = "CTB All Bags"
''  Case Is = UCase("CTBInDsg")
''    Set wfrm = New EmrFrmInvDsg
''    wfrm.Caption = "CTB Design"
''  Case Is = UCase("RepPrm")
''    Set wfrm = New EmrFrpParam
''    wfrm.Caption = "Parameter Listing"
''  Case Is = UCase("RepVarPrm")
''    Set wfrm = New EmrFrpParam
''    wfrm.Caption = "Variable Parameter Listing"
''  Case Is = UCase("RepCustLst")
''    Set wfrm = New EmrFrpCustMst
''    wfrm.Caption = "Customer Listing"
''  Case Is = UCase("RepUsrAccess")
''    Set wfrm = New EmrFrpUsrAccess
''    wfrm.Caption = "User Access Listing"
''  Case Is = UCase("RepCustDsg")
''    Set wfrm = New EmrFrpCustDsg
''    wfrm.Caption = "Customer Design Listing"
''  Case Is = UCase("RepCustAly")
''    Set wfrm = New EmrFrpCustAly
''    wfrm.Caption = "Customer Alloy Listing"
''  Case Is = UCase("RepRmMst")
''    Set wfrm = New EmrFrpRmMst
''    wfrm.Caption = "Raw Material/ Custom Master Listing"
''  Case Is = UCase("RepDsg")
''    Set wfrm = New EmrFrpDsgLst
''    wfrm.Caption = "Design Master Listing"
''  Case Is = UCase("RepDsgCrd")
''    Set wfrm = New EmrFrpDsgCrd
''    wfrm.Caption = "Design Card Details"
''  Case Is = UCase("RepDsgCat")
''    Set wfrm = New EmrFrpDsgCat
''    wfrm.Caption = "Design Catelogue"
''  Case Is = UCase("RepRmRt")
''    Set wfrm = New EmrFrpRmRt
''    wfrm.Caption = "Raw Material Rate Master Listing"
''  Case Is = UCase("RepLabRt")
''    Set wfrm = New EmrFrpLabRt
''    wfrm.Caption = "Labour Rate Master Listing"
''  Case Is = UCase("RepRmRtHist")
''    Set wfrm = New EmrFrpRmRt
''    wfrm.Caption = "Raw Material Rate History Listing"
''  Case Is = UCase("RepLoc")
''    Set wfrm = New EmrFrpLoc
''    wfrm.Caption = "Location Listing"
''  Case Is = UCase("RepDtTbl")
''    Set wfrm = New EmrFrpDtTbl
''    wfrm.Caption = "Date Table Listing"
''  Case Is = UCase("RepBnk")
''    Set wfrm = New EmrFrpBnkAg
''    wfrm.Caption = "Bank Listing"
''  Case Is = UCase("RepBagLst")
''    Set wfrm = New EmrFrpBagLst
''    wfrm.Caption = "Bag Listing"
''  Case Is = UCase("RepBv")
''    Set wfrm = New EmrFrpBv
''    wfrm.Caption = "Bag Movement (Bag Pieces)"
''  Case Is = UCase("RepBagPrn")
''    Set wfrm = New EmrFrpBagPrn
''    wfrm.Caption = "Bag Printing"
''  Case Is = UCase("RepDetOrd")
''    Set wfrm = New EmrFrpOrd
''    wfrm.Caption = "Detail Order Printing"
''  Case Is = UCase("RepAvgOrd")
''    Set wfrm = New EmrFrpOrd
''    wfrm.Caption = "Average Order Printing"
''  Case Is = UCase("RepOrdCat")
''    Set wfrm = New EmrFrpOrd
''    wfrm.Caption = "Order Catalogue Printing"
''  Case Is = UCase("RepOrdProfitMrg")
''    Set wfrm = New EmrFrpOrdProfitMrg
''    wfrm.Caption = "Order Profit Margin"
''  Case Is = UCase("RepRmReqStat")
''    Set wfrm = New EmrFrpRmReq
''    wfrm.Caption = "Raw Material Requirement (Static)"
''  Case Is = UCase("RepRmReqDyn")
''    Set wfrm = New EmrFrpRmReq
''    wfrm.Caption = "Raw Material Requirement (Dynamic)"
''  Case Is = UCase("RepRmReqLoc")
''    Set wfrm = New EmrFrpRmReqLoc
''    wfrm.Caption = "Raw Material Requirement For Loc"
''  Case Is = UCase("RepBagHist")
''    Set wfrm = New EmrFrpBagHist
''    wfrm.Caption = "Bag History"
''  Case Is = UCase("RepPrd")
''    Set wfrm = New EmrFrpPrd
''    wfrm.Caption = "Production Report"
''  Case Is = UCase("RepPrdPts")
''    Set wfrm = New EmrFrpPrdPts
''    wfrm.Caption = "Production Points Report"
''  Case Is = UCase("RepWipBPcs")
''    Set wfrm = New EmrFrpWipBPcs
''    wfrm.Caption = "WIP Bag Pieces Report"
''  Case Is = UCase("RepWipFgBPcs")
''    Set wfrm = New EmrFrpWipBPcs
''    wfrm.Caption = "Fg Bag Pieces Report"
''  Case Is = UCase("RepWipBRm")
''    Set wfrm = New EmrFrpWipBRm
''    wfrm.Caption = "WIP Bag Raw Material Report"
''  Case Is = UCase("RepWipFgBRm")
''    Set wfrm = New EmrFrpWipBRm
''    wfrm.Caption = "Fg Bag Raw Material Report"
''  Case Is = UCase("RepWipBPcsBkDt")
''    Set wfrm = New EmrFrpWipBPcsBkDt
''    wfrm.Caption = "WIP Bag Pieces Report (Back Dated)"
''  Case Is = UCase("RepWipFgBPcsBkDt")
''    Set wfrm = New EmrFrpWipBPcsBkDt
''    wfrm.Caption = "Fg Bag Pieces Report (Back Dated)"
''  Case Is = UCase("RepRejAna")
''    Set wfrm = New EmrFrpRejAna
''    wfrm.Caption = "Rejection Analysis"
''  Case Is = UCase("RepGldLs")
''    Set wfrm = New EmrFrpGldLs
''    wfrm.Caption = "Gold Loss"
''  Case Is = UCase("RepBrkMsgRm")
''    Set wfrm = New EmrFrpBrkMsg
''    wfrm.Caption = "Breakage, Missing Report (Base= Rm)"
''  Case Is = UCase("RepBrkMsgPrd")
''    Set wfrm = New EmrFrpBrkMsg
''    wfrm.Caption = "Breakage, Missing Report (Base= Prd)"
''  Case Is = UCase("RepSalAnaDsg")
''    Set wfrm = New EmrFrpSalAnaDsg
''    wfrm.Caption = "Sales Analysis (Design Level)"
''  Case Is = UCase("RepSalAnaRm")
''    Set wfrm = New EmrFrpSalAnaRm
''    wfrm.Caption = "Sales Analysis (Rm Level)"
''  Case Is = UCase("RepSalAnaLab")
''    Set wfrm = New EmrFrpSalAnaLab
''    wfrm.Caption = "Sales Analysis (Lab Level)"
''  Case Is = UCase("RepDiaProfit")
''    Set wfrm = New EmrFrpDiaProfit
''    wfrm.Caption = "Diamond Profit"
''  Case Is = UCase("RepSOPnd")
''    Set wfrm = New EmrFrpSOPnd
''    wfrm.Caption = "Sales Order Pending"
''  Case Is = UCase("RepWIWPrd")
''    Set wfrm = New EmrFrpWIW
''    wfrm.Caption = "What Is Where (Production)"
''  Case Is = UCase("RepWIWExp")
''    Set wfrm = New EmrFrpWIW
''    wfrm.Caption = "What Is Where (Export)"
''  Case Is = UCase("RepWIWLongPrd")
''    Set wfrm = New EmrFrpWIW
''    wfrm.Caption = "What Is Where Long (Production)"
''  Case Is = UCase("RepWIWLongExp")
''    Set wfrm = New EmrFrpWIW
''    wfrm.Caption = "What Is Where Long (Export)"
''  Case Is = UCase("RepStkLed")
''    Set wfrm = New EmrFrpStkLed
''    wfrm.Caption = "Stock Ledger"
''  Case Is = UCase("RepFlashStk")
''    Set wfrm = New EmrFrpFlashStk
''    wfrm.Caption = "Flash Stock"
''  Case Is = UCase("RepCustmStk")
''    Set wfrm = New EmrFrpCustmStk
''    wfrm.Caption = "Custom Stock"
''  Case Is = UCase("RepStkSumm")
''    Set wfrm = New EmrFrpStkSumm
''    wfrm.Caption = "Stock Summary"
''  Case Is = UCase("RepTxnDT")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Detail Transaction Listing"
''  Case Is = UCase("RepTxnBS")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Bag Split Transaction Listing"
''  Case Is = UCase("RepTxnMLT")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Melting Transaction Listing"
''  Case Is = UCase("RepTxnBV")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Bag Movement Transaction Listing"
''  Case Is = UCase("RepTxnREJ")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Rejection Transaction Listing"
''  Case Is = UCase("RepTxnPR")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Purchase/Sales Transaction Listing"
''  Case Is = UCase("RepTxnMV")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "RM Movement Transaction Listing"
''  Case Is = UCase("RepTxnCNV")
''    Set wfrm = New EmrFrpTxnLst
''    wfrm.Caption = "Conversion Transaction Listing"
''  Case Is = UCase("RepUnMtchCnv")
''    Set wfrm = New EmrFrpUnMtchCnv
''    wfrm.Caption = "Unmatched Conversion Listing"
''  Case Is = UCase("RepFbLst")
''    Set wfrm = New EmrFrpFgList
''    wfrm.Caption = "Finished Goods (Bag)"
''  Case Is = UCase("RepFwbLst")
''    Set wfrm = New EmrFrpFgList
''    wfrm.Caption = "Finished Goods (Bag Return)"
''  Case Is = UCase("RepFgMvm")
''    Set wfrm = New EmrFrpFgMvm
''    wfrm.Caption = "Finished Goods (Bag Movement)"
''  Case Is = UCase("RepVchPrnDT")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Daily Voucher Printing"
''  Case Is = UCase("RepVchPrnBS")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Bag Split Voucher Printing"
''  Case Is = UCase("RepVchPrnMLT")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Melting Voucher Printing"
''  Case Is = UCase("RepVchPrnBV")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Bag Movement Voucher Printing"
''  Case Is = UCase("RepVchPrnREJ")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Rejection Voucher Printing"
''  Case Is = UCase("RepVchPrnPR")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Purchase/Sales Voucher Printing"
''  Case Is = UCase("RepVchPrnMV")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "RM Movement Voucher Printing"
''  Case Is = UCase("RepVchPrnCNV")
''    Set wfrm = New EmrFrpVchPrn
''    wfrm.Caption = "Conversion Voucher Printing"
''  Case Is = UCase("RepDetInv")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "Invoice Costing"
''  Case Is = UCase("RepAvgInv")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "Average Invoice Printing"
''  Case Is = UCase("RepInvCat")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "Invoice Catalogue"
''  Case Is = UCase("RepInv")
''    Set wfrm = New EmrFrpInv
''    wfrm.Caption = "Invoice Printing"
''  Case Is = UCase("RepGrFormFrnt")
''    Set wfrm = New EmrFrpInv
''    wfrm.Caption = "GR Form (Front)"
''  Case Is = UCase("RepGrFormBack")
''    Set wfrm = New EmrFrpInvLet
''    wfrm.Caption = "GR Form (Back)"
''  Case Is = UCase("RepShpBill")
''    Set wfrm = New EmrFrpInv
''    wfrm.Caption = "Shipping Bill"
''  Case Is = UCase("RepPckLst")
''    Set wfrm = New EmrFrpPckLst
''    wfrm.Caption = "Finished Goods Packing List Fr INV"
''  Case Is = UCase("RepPckLstInv")
''    Set wfrm = New EmrFrpPckLstInv
''    wfrm.Caption = "Invoice Packing List"
''  Case Is = UCase("RepInvBOELst")
''    Set wfrm = New EmrFrpInvBoeLst
''    wfrm.Caption = "Invoice BOE Listing"
''  Case Is = UCase("RepGSPFORM")
''    Set wfrm = New EmrFrpInv
''    wfrm.Caption = "GSP Form"
''  Case Is = UCase("RepGSPDetLet")
''    Set wfrm = New EmrFrpInvLet
''    wfrm.Caption = "GSP Detail Letter"
''  Case Is = UCase("RepGSPCert")
''    Set wfrm = New EmrFrpInvLet
''    wfrm.Caption = "GSP Certification"
''  Case Is = UCase("RepDeclnLet")
''    Set wfrm = New EmrFrpInvLet
''    wfrm.Caption = "Declaration Letter"
''  Case Is = UCase("RepTagPrn")
''    'Dim wfrm As New EmrFrpTagPrn
''    Set wfrm = New EmrFrmTag
''    wfrm.Caption = "Tag Printing"
''  Case Is = UCase("RepInvProfitMrg")
''    Set wfrm = New EmrFrpInvProfitMrg
''    wfrm.Caption = "Profit Margin"
''  Case Is = UCase("RepInvLossSubRep")
''    Set wfrm = New EmrFrpInvLossSubRep
''    wfrm.Caption = "Invoice Metal Loss Sub Report"
''  Case Is = UCase("RepAnnexValAddn")
''    Set wfrm = New EmrFrpValAddnAnnex
''    wfrm.Caption = "Annexure To Value Addition"
''  Case Is = UCase("RepValAddn")
''    Set wfrm = New EmrFrpValAddn
''    wfrm.Caption = "Value Addition"
''  Case Is = UCase("RepDetInvCTB")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "CTB Invoice Costing"
''  Case Is = UCase("RepAvgInvCTB")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "Average CTB Invoice Printing"
''  Case Is = UCase("RepInvCTBCat")
''    Set wfrm = New EmrFrpInvCost
''    wfrm.Caption = "CTB Invoice Catalogue"
''  Case Is = UCase("RepCTBPckLst")
''    Set wfrm = New EmrFrpPckLst
''    wfrm.Caption = "Finished Goods Packing List For CTB"
''  Case Is = UCase("PpcCap")
''    Set wfrm = New EmrFrmPpcCap
''  Case Is = UCase("PpcDsgLd")
''    Set wfrm = New EmrFrmPpcDsgLd
''  Case Is = UCase("PpcOrdPln")
''    Set wfrm = New EmrFrmPpcOrdPln
''  Case Is = UCase("RepDsgLd")
''    Set wfrm = New EmrFrpPpcDsgLd
''  Case Is = UCase("RepCntrlCpy")
''    Set wfrm = New EmrFrpPpcCntrlCpy
''  Case Is = UCase("RepOrdPln")
''    Set wfrm = New EmrFrpPpcOrdPln
''  Case Is = UCase("XMLMstToEmr")
''    Set wfrm = New xEmrFrmTrfIn
''    wfrm.Caption = "XML To Emr (Master)"
''  Case Is = UCase("XMLToEmr")
''    Set wfrm = New EmrFrmTrfIn
''    wfrm.Caption = "XML To Emr (Order)"
''  Case Is = UCase("Zoom")
''    Set wfrm = New EmrFrmOrdZoom
''  Case Is = UCase("LsRec")
''    Set wfrm = New EmrFrmLsRec
''  Case Is = UCase("RepLsRec")
''    Set wfrm = New EmrFrpLsRec
''  Case Is = UCase("RepCustRm")
''    Set wfrm = New EmrFrpCustRm
'  End Select
'
'  If Not (wfrm Is Nothing) Then
'    ShowForm wfrm, ps_Mnu
'  End If
'  '*** Jenny Color
'End Sub

' Zubin 213
Private Function ResolutionChk() As Boolean
  Dim xTwips As Integer, yTwips As Integer
  Dim xPixels As Integer, yPixels As Integer
  
  xTwips = Screen.TwipsPerPixelX
  yTwips = Screen.TwipsPerPixelY

  xPixels = Screen.Width / xTwips
  yPixels = Screen.Height / yTwips
  
  'MsgBox CStr(xPixels) + " x " + CStr(yPixels)
  If xPixels < 1024 Or yPixels < 768 Then
    ResolutionChk = False
    Exit Function
  End If
  ResolutionChk = True
End Function

Private Sub OldFillSrvrCmb()
    '********** 3.11.4 *******************
    Dim wRs As MwfLib.MDORowSet
    'Dim wo_MstConn As MwfLib.MDOConnection
    'Set wo_MstConn = Ctl.GetMwMstConn(ctProjectName)
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If
   cmbSrvrNm.Clear
   Set wRs = mo_MstConn.OpenRes(" Select MsSrvrNm from MSys Where " + _
                                       " MsLoginNm='D6CED2D13EDC' And MsPwd='D6CED2D13EDC36228E' ")
   Dim i As Integer
   If Not (wRs.EOF Or wRs.BOF) Then
    
    For i = 1 To wRs.RecCount
        cmbSrvrNm.AddItem wRs!MsSrvrNm
        wRs.MoveNext
    Next
    If cmbSrvrNm.ListCount > 0 Then cmbSrvrNm.ListIndex = 0
   End If
   
   Call FillDbCmb      '3.11.4
   Set wRs = Nothing
   'Set wo_MstConn = Nothing
   '********** 3.11.4 *******************
End Sub

Private Sub OldFillDbCmb()
    '********** 3.11.4 *******************
    Dim wRs As MwfLib.MDORowSet
'    Dim wo_MstConn As MwfLib.MDOConnection
'    Set wo_MstConn = MWLib.GetMwMstConn(ctProjectName)
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If
    cmbDbNm.Clear
   Set wRs = mo_MstConn.OpenRes(" Select MsLoginNm from MSys where " + _
                                       " MsSrvrNm='EMR' and MsDbSrvrNm='" + cmbSrvrNm.text + "'")

   Dim i As Integer, wDbNm As String, wLstIdx As Integer
   If Not (wRs.EOF Or wRs.BOF) Then
    wDbNm = mo_MstConn.GetFldVal("Select Top 1 MuDbNm From MUsrMst Where MuSrvrNm='" + cmbSrvrNm.text + "' And MuSysId='" + ctProjectName + "' And MuUsrId In ('','" + TXT_UCD + "') Order by MuUsrId Desc ")
        For i = 1 To wRs.RecCount
            cmbDbNm.AddItem wRs!MsLoginNm
            If UCase(wRs!MsLoginNm) = UCase(wDbNm) Then wLstIdx = i
            wRs.MoveNext
        Next
    If cmbDbNm.ListCount > 0 Then cmbDbNm.ListIndex = IIF(wLstIdx > 0, wLstIdx - 1, 0)
   End If

   Set wRs = Nothing
   'Set wo_MstConn = Nothing
   '********** 3.11.4 *******************
End Sub
Private Sub FillSrvrCmb()
    '********** 3.11.4 *******************
    Dim wRs As MwfLib.MDORowSet
     If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If
   cmbSrvrNm.Clear
   Set wRs = mo_MstConn.OpenRes(" Select MsSrvrNm from MEmrSys Where " + _
                                       " MsSysId = '" + ctProjectName + "'")

   Dim i As Integer
   If Not (wRs.EOF Or wRs.BOF) Then
    For i = 1 To wRs.RecCount
        cmbSrvrNm.AddItem wRs!MsSrvrNm
        wRs.MoveNext
    Next
    If cmbSrvrNm.ListCount > 0 Then cmbSrvrNm.ListIndex = 0
   End If
   
   Call FillDbCmb      '3.11.4
   Set wRs = Nothing
   'Set wo_MstConn = Nothing
   '********** 3.11.4 *******************
End Sub

Private Sub FillDbCmb()
    '********** 3.11.4 *******************
    Dim wRs As MwfLib.MDORowSet
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If
     cmbDbNm.Clear
     Set wRs = mo_MstConn.OpenRes(" Select MsDbNm from MEmrSys where MsSysId = '" + ctProjectName + "'" + _
                                       "  and  MsSrvrNm ='" + cmbSrvrNm.text + "'")

   Dim i As Integer, wDbNm As String, wLstIdx As Integer
   If Not (wRs.EOF Or wRs.BOF) Then
        wDbNm = mo_MstConn.GetFldVal("Select MuDbNm From MUsrMst Where MuSysId='" + ctProjectName + "' And MuUsrId = '" + TXT_UCD + "'")
        For i = 1 To wRs.RecCount
            cmbDbNm.AddItem wRs!MsDbNm
            If UCase(wRs!MsDbNm) = UCase(wDbNm) Then wLstIdx = i
            wRs.MoveNext
        Next
    If cmbDbNm.ListCount > 0 Then cmbDbNm.ListIndex = IIF(wLstIdx > 0, wLstIdx - 1, 0)
   End If
   Set wRs = Nothing
End Sub

