VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "*\A..\61CtlBack\MwfCtl.vbp"
Begin VB.MDIForm EmrFrmMDI 
   BackColor       =   &H00808080&
   Caption         =   "Emperor"
   ClientHeight    =   3030
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   5145
   Icon            =   "emrfrmmdi.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "emrfrmmdi.frx":1272
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   0
      ScaleHeight     =   2985
      ScaleWidth      =   5145
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   5145
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
      Begin MSCommLib.MSComm MSComm1 
         Left            =   90
         Top             =   870
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
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
         Left            =   5190
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "C"
         Top             =   2460
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
         Left            =   4560
         Top             =   0
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
         Pic             =   "emrfrmmdi.frx":3D3C1
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
      Begin VB.PictureBox MWCTL_HLP1 
         Height          =   285
         Left            =   6120
         ScaleHeight     =   225
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
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
         Pic             =   "emrfrmmdi.frx":3EAD3
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
         Left            =   810
         TabIndex        =   22
         Top             =   1860
         Width           =   1455
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
         Left            =   810
         TabIndex        =   21
         Top             =   1537
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
         Left            =   3690
         TabIndex        =   20
         Top             =   2490
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
         Left            =   4320
         Picture         =   "emrfrmmdi.frx":401E5
         Top             =   120
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
      Top             =   2985
      Visible         =   0   'False
      Width           =   5145
      _ExtentX        =   9075
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
      Begin VB.OLE OLE1 
         Class           =   "AcroExch.Document.7"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   23
         Top             =   -120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   19
      Top             =   2745
      Visible         =   0   'False
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6985
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
            Object.Width           =   1058
            MinWidth        =   1058
            Key             =   "DiskSpace"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   159
            MinWidth        =   35
            Key             =   "ErrPct"
            Object.ToolTipText     =   "Click to View Previous Error"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Begin VB.Menu Hlp 
         Caption         =   "Help"
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
         Begin VB.Menu ContractorMst 
            Caption         =   "Contractor Master"
            Tag             =   "215"
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
         Begin VB.Menu PRDPTS 
            Caption         =   "Production Points Master"
            Tag             =   "253"
         End
         Begin VB.Menu LossEntry 
            Caption         =   "Job Work Loss Entry"
            Tag             =   "254"
         End
         Begin VB.Menu Bnk 
            Caption         =   "Bank/Agent Master"
            Tag             =   "255"
         End
         Begin VB.Menu RfIdMst 
            Caption         =   "Bag Alias Master"
            Tag             =   "256"
         End
         Begin VB.Menu PrcLMst 
            Caption         =   "Process Loss Master"
            Tag             =   "257"
         End
         Begin VB.Menu PRDPTSSKETCH 
            Caption         =   "Production Points (Sketch)"
            Tag             =   "258"
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
   Begin VB.Menu TCS 
      Caption         =   "&Transactions"
      Tag             =   "3"
      Visible         =   0   'False
      Begin VB.Menu Bag 
         Caption         =   "Bag Opening"
         Tag             =   "31"
      End
      Begin VB.Menu Ord 
         Caption         =   "Order/Quotation/PDC"
         Tag             =   "32"
         Begin VB.Menu SalOrd 
            Caption         =   "Sales Order / Work Order Entry"
            Tag             =   "321"
         End
         Begin VB.Menu SalOrdAck 
            Caption         =   "Sales Order Acknowledgement"
            Tag             =   "326"
         End
         Begin VB.Menu PdcmMnu 
            Caption         =   "PDCM"
            Tag             =   "322"
            Begin VB.Menu PrcLst 
               Caption         =   "Party-Dsg Combination Master"
               Tag             =   "3221"
            End
            Begin VB.Menu PdcmUpd 
               Caption         =   "Auto Update PDCM"
               Tag             =   "3222"
            End
            Begin VB.Menu DPdcm 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu Quot 
            Caption         =   "Quotation Entry"
            Tag             =   "323"
         End
         Begin VB.Menu QB 
            Caption         =   "Quotation For FG Bag"
            Tag             =   "324"
         End
         Begin VB.Menu SkQuot 
            Caption         =   "Sketch Quotation Entry"
            Tag             =   "325"
         End
         Begin VB.Menu MultiPrcQt 
            Caption         =   "Multi Price Quotation"
            Tag             =   "327"
         End
         Begin VB.Menu SalOrdSk 
            Caption         =   "Sale Order Sketch"
            Tag             =   "328"
         End
         Begin VB.Menu D31 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu BagTxn 
         Caption         =   "Bag Transactions"
         Tag             =   "33"
         Visible         =   0   'False
         Begin VB.Menu Dt 
            Caption         =   "Daily Transaction"
            Tag             =   "331"
         End
         Begin VB.Menu BV 
            Caption         =   "Bag Movement"
            Tag             =   "332"
         End
         Begin VB.Menu BA 
            Caption         =   "Bag Acknowledgement"
            Tag             =   "333"
         End
         Begin VB.Menu BS 
            Caption         =   "Bag Split/Bag Transfer"
            Tag             =   "334"
         End
         Begin VB.Menu BM 
            Caption         =   "Bag Merge"
            Tag             =   "335"
         End
         Begin VB.Menu REJ 
            Caption         =   "Rejection"
            Tag             =   "336"
         End
         Begin VB.Menu MLT 
            Caption         =   "Melting"
            Tag             =   "337"
         End
         Begin VB.Menu D32 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RmTxn 
         Caption         =   "Raw Material Transactions"
         Tag             =   "34"
         Visible         =   0   'False
         Begin VB.Menu PR 
            Caption         =   "Purchase/Receipt"
            Tag             =   "341"
         End
         Begin VB.Menu Sa 
            Caption         =   "Sales"
            Tag             =   "342"
         End
         Begin VB.Menu TZ 
            Caption         =   "Customs Purchase"
            Tag             =   "343"
         End
         Begin VB.Menu MV 
            Caption         =   "Raw Material Issue"
            Tag             =   "344"
         End
         Begin VB.Menu CNV 
            Caption         =   "Conversion"
            Tag             =   "345"
         End
         Begin VB.Menu MA 
            Caption         =   "Acknowledgement"
            Tag             =   "346"
         End
         Begin VB.Menu D33 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu FG 
         Caption         =   "Finished Goods Transaction"
         Tag             =   "35"
         Begin VB.Menu FB 
            Caption         =   "Finished Goods (Bag)"
            Tag             =   "351"
         End
         Begin VB.Menu FWB 
            Caption         =   "Finished Goods (Bag Return)"
            Tag             =   "352"
         End
         Begin VB.Menu FM 
            Caption         =   "Finished Goods (Bag Movement)"
            Tag             =   "353"
         End
         Begin VB.Menu D34 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu INV 
         Caption         =   "Invoice Entries"
         Tag             =   "36"
         Begin VB.Menu ShpNote 
            Caption         =   "Shipment Note"
            Tag             =   "365"
         End
         Begin VB.Menu IN 
            Caption         =   "Invoice Header"
            Tag             =   "361"
         End
         Begin VB.Menu InAllBags 
            Caption         =   "Invoice All Bags"
            Tag             =   "362"
         End
         Begin VB.Menu InDsg 
            Caption         =   "Invoice Design"
            Tag             =   "363"
         End
         Begin VB.Menu InCustmMtch 
            Caption         =   "Invoice Custom Matching"
            Tag             =   "364"
         End
         Begin VB.Menu D35 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RetMem 
         Caption         =   "Return Memo"
         Tag             =   "3A"
      End
      Begin VB.Menu D3 
         Caption         =   "DUM"
      End
      Begin VB.Menu Tree 
         Caption         =   "Tree Entry"
         Tag             =   "39"
      End
   End
   Begin VB.Menu Rep 
      Caption         =   "&Reports"
      Tag             =   "4"
      Visible         =   0   'False
      Begin VB.Menu RepDaily 
         Caption         =   "Daily Reports"
         Tag             =   "41"
         Begin VB.Menu RepWIW 
            Caption         =   "What Is Where"
            Tag             =   "411"
            Begin VB.Menu RepWIWPrd 
               Caption         =   "What Is Where (Prd)"
               Tag             =   "4111"
            End
            Begin VB.Menu RepWIWExp 
               Caption         =   "What Is Where (Exp)"
               Tag             =   "4112"
            End
            Begin VB.Menu RepWIWLongPrd 
               Caption         =   "What Is Where Long (Prd)"
               Tag             =   "4113"
            End
            Begin VB.Menu RepWIWLongExp 
               Caption         =   "What Is Where Long (Exp)"
               Tag             =   "4114"
            End
            Begin VB.Menu D411 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepWIP 
            Caption         =   "WIP"
            Tag             =   "412"
            Begin VB.Menu RepWipBPcs 
               Caption         =   "WIP Bag Pieces"
               Tag             =   "4121"
            End
            Begin VB.Menu RepWipBRm 
               Caption         =   "WIP Bag Raw Material"
               Tag             =   "4122"
            End
            Begin VB.Menu RepWipBPcsBkDt 
               Caption         =   "WIP Bag Pieces (Back Dated)"
               Tag             =   "4123"
            End
            Begin VB.Menu RepWipBRmBkDt 
               Caption         =   "WIP Bag RM (Back Dated)"
               Tag             =   "4129"
            End
            Begin VB.Menu RepWipFgBPcs 
               Caption         =   "Fg Bag Pieces"
               Tag             =   "4124"
            End
            Begin VB.Menu RepWipFgBRm 
               Caption         =   "Fg Bag Raw Material"
               Tag             =   "4125"
            End
            Begin VB.Menu RepWipFgBPcsBkDt 
               Caption         =   "Fg Bag Pieces (Back Dated)"
               Tag             =   "4126"
            End
            Begin VB.Menu RepTagPrnFg 
               Caption         =   "FG Tag Printing"
               Tag             =   "4127"
            End
            Begin VB.Menu RepTagPrnBagFg 
               Caption         =   "Bag FG Tag Printing"
               Tag             =   "4128"
            End
            Begin VB.Menu D412 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepPrd 
            Caption         =   "Production Report"
            Tag             =   "413"
         End
         Begin VB.Menu REPBM 
            Caption         =   "Bag Movement"
            Tag             =   "414"
            Begin VB.Menu RepBv 
               Caption         =   "Bag Movement Bag Pieces"
               Tag             =   "4141"
            End
            Begin VB.Menu RepBvRm 
               Caption         =   "Bag Movement Bag Rm"
               Tag             =   "4142"
            End
            Begin VB.Menu D10 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepBagPrnMnu 
            Caption         =   "Bag Printing"
            Tag             =   "415"
            Begin VB.Menu RepBagPrn 
               Caption         =   "Primary Bag Printing "
               Tag             =   "4151"
            End
            Begin VB.Menu RepFltBagPrn 
               Caption         =   "Flute Bag Printing"
               Tag             =   "4152"
            End
            Begin VB.Menu D9 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepOrd 
            Caption         =   "Order"
            Tag             =   "417"
            Begin VB.Menu RepDetOrd 
               Caption         =   "Detail Order Printing"
               Tag             =   "4171"
            End
            Begin VB.Menu RepAvgOrd 
               Caption         =   "Average Order Printing"
               Tag             =   "4172"
            End
            Begin VB.Menu RepOrdCat 
               Caption         =   "Order Catalogue"
               Tag             =   "4173"
            End
            Begin VB.Menu RepOrdProfitMrg 
               Caption         =   "Profit Margin"
               Tag             =   "4174"
            End
            Begin VB.Menu Zoom 
               Caption         =   "Zoom"
               Tag             =   "4175"
            End
            Begin VB.Menu RepMultiPrcQt 
               Caption         =   "Multi Price Quotation Printing"
               Tag             =   "4177"
            End
            Begin VB.Menu D413 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepStk 
            Caption         =   "Stock"
            Tag             =   "418"
            Begin VB.Menu RepStkLed 
               Caption         =   "Stock Ledger"
               Tag             =   "4181"
            End
            Begin VB.Menu RepFlashStk 
               Caption         =   "Flash Stock"
               Tag             =   "4182"
            End
            Begin VB.Menu RepCustmStk 
               Caption         =   "Custom Stock"
               Tag             =   "4183"
            End
            Begin VB.Menu RepCustmStkLand 
               Caption         =   "Custom Stock (Landscape)"
               Tag             =   "4185"
            End
            Begin VB.Menu RepFgBagLed 
               Caption         =   "Finished Goods Bag Ledger"
               Tag             =   "4184"
            End
            Begin VB.Menu D414 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu D41 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepAna 
         Caption         =   "Analytical Reports"
         Tag             =   "42"
         Begin VB.Menu RepRmReq 
            Caption         =   "Raw Material Requirement"
            Tag             =   "421"
            Begin VB.Menu RepRmReqStat 
               Caption         =   "Static"
               Tag             =   "4211"
            End
            Begin VB.Menu RepRmReqDyn 
               Caption         =   "Dynamic"
               Tag             =   "4212"
            End
            Begin VB.Menu RepRmReqLoc 
               Caption         =   "Location"
               Tag             =   "4213"
            End
            Begin VB.Menu REPNETSTK 
               Caption         =   "Net Stock"
               Tag             =   "4214"
            End
            Begin VB.Menu RepRmReqSet 
               Caption         =   "Setting Load"
               Tag             =   "4215"
            End
            Begin VB.Menu D421 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepBagHist 
            Caption         =   "Bag History"
            Tag             =   "422"
         End
         Begin VB.Menu RepRejAna 
            Caption         =   "Rejection Analysis"
            Tag             =   "423"
         End
         Begin VB.Menu RepMetLs 
            Caption         =   "Metal Loss"
            Tag             =   "424"
            Begin VB.Menu RepMetLsWip 
               Caption         =   "Metal Loss (WIP)"
               Tag             =   "4241"
            End
            Begin VB.Menu RepMetlsFg 
               Caption         =   "Metal Loss (FG)"
               Tag             =   "4242"
            End
            Begin VB.Menu D424 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepBrkMsg 
            Caption         =   "Breakage, Missing Report"
            Tag             =   "425"
            Begin VB.Menu RepBrkMsgRm 
               Caption         =   "Analysis Based On Rm"
               Tag             =   "4251"
            End
            Begin VB.Menu RepBrkMsgPrd 
               Caption         =   "Analysis Based On Prd"
               Tag             =   "4252"
            End
            Begin VB.Menu D422 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepSalAna 
            Caption         =   "Sales Analysis"
            Tag             =   "426"
            Begin VB.Menu RepSalAnaDsg 
               Caption         =   "Design Level"
               Tag             =   "4261"
            End
            Begin VB.Menu RepSalAnaRm 
               Caption         =   "Raw Material Level"
               Tag             =   "4262"
            End
            Begin VB.Menu RepSalAnaLab 
               Caption         =   "Labour Level"
               Tag             =   "4263"
            End
            Begin VB.Menu D423 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepDiaProfit 
            Caption         =   "Diamond Profit"
            Tag             =   "427"
         End
         Begin VB.Menu RepSOPnd 
            Caption         =   "Sales Order Pending"
            Tag             =   "428"
         End
         Begin VB.Menu REPSOPNDRM 
            Caption         =   "SO Pending RM Detail"
            Tag             =   "42A"
         End
         Begin VB.Menu RepShpAna 
            Caption         =   "Shipment Analysis"
            Tag             =   "42B"
         End
         Begin VB.Menu RepStkSumm 
            Caption         =   "Stock Summary"
            Tag             =   "429"
         End
         Begin VB.Menu D42 
            Caption         =   "DUM"
         End
      End
      Begin VB.Menu RepIn 
         Caption         =   "Invoice"
         Tag             =   "44"
         Begin VB.Menu RepInvCst 
            Caption         =   "Costing Reports"
            Tag             =   "441"
            Begin VB.Menu eINVGEN 
               Caption         =   "eInvoice Generation"
            End
            Begin VB.Menu RepDetInv 
               Caption         =   "Detail Invoice Costing"
               Tag             =   "4411"
            End
            Begin VB.Menu RepAvgInv 
               Caption         =   "Average Invoice Costing"
               Tag             =   "4412"
            End
            Begin VB.Menu RepInvCat 
               Caption         =   "Invoice Catalogue"
               Tag             =   "4413"
            End
            Begin VB.Menu RepInvCustStmt 
               Caption         =   "Customer Statement"
               Tag             =   "4414"
            End
            Begin VB.Menu D441 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepInvStat 
            Caption         =   "Statutory Reports"
            Tag             =   "442"
            Begin VB.Menu RepShpBill 
               Caption         =   "Shipping Bill"
               Tag             =   "4421"
            End
            Begin VB.Menu RepGrFormFrnt 
               Caption         =   "GR Form (Front)"
               Tag             =   "4422"
            End
            Begin VB.Menu RepGrFormBack 
               Caption         =   "GR Form (Back)"
               Tag             =   "4423"
            End
            Begin VB.Menu RepGSPFORM 
               Caption         =   "GSP Form"
               Tag             =   "4424"
            End
            Begin VB.Menu RepGSPDetLet 
               Caption         =   "GSP Detail Letter"
               Tag             =   "4425"
            End
            Begin VB.Menu RepGSPCert 
               Caption         =   "GSP Certification"
               Tag             =   "4426"
            End
            Begin VB.Menu RepDeclnLet 
               Caption         =   "Declaration Letter"
               Tag             =   "4427"
            End
            Begin VB.Menu RepShpBillMulti 
               Caption         =   "Shipping Bill (Multi Inv)"
               Tag             =   "4428"
            End
            Begin VB.Menu RepGrFormFrntMulti 
               Caption         =   "GR Form (Front) (Multi Inv)"
               Tag             =   "4429"
            End
            Begin VB.Menu RepGRFormBackMulti 
               Caption         =   "GR Form (Back) (Multi Inv)"
               Tag             =   "442A"
            End
            Begin VB.Menu RepGSPFORMMulti 
               Caption         =   "GSP Form (Multi)"
               Tag             =   "442B"
            End
            Begin VB.Menu RepGSPDetLetMulti 
               Caption         =   "GSP Detail Letter (Multi Inv)"
               Tag             =   "442C"
            End
            Begin VB.Menu RepGSPCertMulti 
               Caption         =   "GSP Certification (Multi Inv)"
               Tag             =   "442D"
            End
            Begin VB.Menu RepDeclnLetMulti 
               Caption         =   "Declaration Letter (Multi Inv)"
               Tag             =   "442E"
            End
            Begin VB.Menu RepShpBillFile 
               Caption         =   "Shipping Bill File"
               Tag             =   "442H"
            End
            Begin VB.Menu D442 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepInvOth 
            Caption         =   "Other Reports"
            Tag             =   "443"
            Begin VB.Menu RepInv 
               Caption         =   "Invoice Printing"
               Tag             =   "4431"
            End
            Begin VB.Menu RepPckLst 
               Caption         =   "Packing List (Finished Goods)"
               Tag             =   "4432"
            End
            Begin VB.Menu RepPckLstInv 
               Caption         =   "Packing List (Invoice)"
               Tag             =   "4433"
            End
            Begin VB.Menu RepPckLstInvBag 
               Caption         =   "Packing List (Invoice Bag)"
               Tag             =   "4439"
            End
            Begin VB.Menu RepTagPrn 
               Caption         =   "Invoice Tag Printing"
               Tag             =   "4434"
            End
            Begin VB.Menu RepInvProfitMrg 
               Caption         =   "Invoice Profit Margin"
               Tag             =   "4435"
            End
            Begin VB.Menu RepInvBOELst 
               Caption         =   "Invoice BOE Listing"
               Tag             =   "4436"
            End
            Begin VB.Menu RepInvLossSubRep 
               Caption         =   "Invoice Metal Loss Summary"
               Tag             =   "4437"
            End
            Begin VB.Menu RepLclInv 
               Caption         =   "Local Invoice"
               Tag             =   "4438"
            End
            Begin VB.Menu D443 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepInvValAdd 
            Caption         =   "Value Addition"
            Tag             =   "444"
            Begin VB.Menu RepValAddn 
               Caption         =   "Value Addition"
               Tag             =   "4441"
            End
            Begin VB.Menu RepAnnexValAddn 
               Caption         =   "Annexure To Value Addition"
               Tag             =   "4442"
            End
            Begin VB.Menu RepValAddnDta 
               Caption         =   "Value Addition (DTA)"
               Tag             =   "4443"
            End
            Begin VB.Menu D444 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu RepInLoose 
            Caption         =   "Loose RM Export"
            Tag             =   "445"
            Begin VB.Menu RepInvLoose 
               Caption         =   "Invoice Printing"
               Tag             =   "4451"
            End
            Begin VB.Menu RepShpBillLoose 
               Caption         =   "Shipping Bill"
               Tag             =   "4452"
            End
            Begin VB.Menu RepGrFormLoose 
               Caption         =   "GR Fom"
               Tag             =   "4453"
            End
            Begin VB.Menu RepPckLstLoose 
               Caption         =   "Packing List"
               Tag             =   "4454"
            End
            Begin VB.Menu D445 
               Caption         =   "DUM"
            End
         End
         Begin VB.Menu D44 
            Caption         =   "DUM"
         End
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
      Begin VB.Menu TRACK 
         Caption         =   "Track Status"
      End
      Begin VB.Menu GENMNU 
         Caption         =   "Generate Menu"
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
Attribute VB_Name = "EmrFrmMDI"
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


Private Sub ShpNote_Click()
  Call FormSelect("ShpNote")
End Sub

Private Sub CertEntry_Click()
'VK.5-added menu for Certificate Entry
  Call FormSelect("CertEntry")
End Sub
Private Sub eInv_Click()
  Call FormSelect("eInv")
End Sub

Private Sub CertReport_Click()
  Call FormSelect("CertReport")
End Sub
Private Sub cmbSrvrNm_Click()
Call FillDbCmb      '3.11.4
End Sub
Private Sub PrcLMst_Click()
  '3.11.0
  Call FormSelect("PrcLMst")
End Sub
'Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As ComctlLib.ButtonMenu)
'  '*** Jen ***
'  'Place a toolBar on a cool bar. Add Buttons to the ToolBar. Add ButtonMenus to the button.
'  'If need be add Popup Menus also.
'  'In the general tab of the ToolBar properties set the TextAlignment as '1 - tbrTextAlignRight'
'  'In the buttons tab, keep the Style of the button as '5 - tbrDropdown'
'  'If a popup menu has to be displayed for a but tonmenu then a key has to be specified
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

Private Sub CTB_Click()

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

Private Sub DsgPrmCat_Click()
'*** Jay 3.4[DP]
  Call FormSelect("DsgPrmCat")
End Sub

' ***** Manali 3.9.1 - 21/02/12 - Color Mnus shifted to AddonMnus
Private Sub ClrDfn_Click()
  Call FormSelect("CLRDFN")
End Sub
' ***** Manali 3.9.1 - 21/02/12 - Color Mnus shifted to AddonMnus
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

Private Sub OpenPort()
On Error GoTo errhld
  Dim wRs_Port As MDORowSet
  Set wRs_Port = ADC.Connection.OpenRes _
                ("Select * From Param where PTyp='METLR' and " + _
                "(','+PDesc225+',' like '%," + gs_UsrCd + ",%' or (Select HMetlrCd From head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "')=PMCd) " + _
                "Order By (Case When ','+PDesc225+',' like '%," + gs_UsrCd + ",%' Then 1 Else 2 End)")
  If Not (wRs_Port.EOF Or wRs_Port.BOF) Then
    MSComm1.CommPort = wRs_Port!PNum
    MSComm1.Settings = wRs_Port!pValue
    MSComm1.RThreshold = wRs_Port!PNum1             ' RThreshold works with values 1 or 10 +
    MSComm1.InputLen = 0                            ' This means max input length allowed
    MSComm1.InputMode = comInputModeText            ' comInputModeBinary does not work
    MSComm1.PortOpen = True
    wRs_Port.MoveNext
  End If
  Exit Sub
errhld:
  DispMsg Err.Description, etError
End Sub

Private Sub ClosePort()
On Error GoTo errhld
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Exit Sub
errhld:
    DispMsg Err.Description, etError
End Sub

Private Sub MSComm1_OnComm()
On Error GoTo errhld
  Dim wInPut As String, wInWt As Double
  
  If MSComm1.CommEvent = comEvReceive Then
    wInPut = MSComm1.Input
    
    wInPut = Trim(CStr(wInPut))
    wInWt = GetNumVal(Trim(wInPut))
    
    With Me.ActiveForm
    Select Case UCase(.Name)
    
    Case Is = UCase("EmrFrmTxn")
        Select Case UCase(.ActiveControl.Name)
        Case Is = UCase("ATXT")
            If UCase(.ActiveControl.IdName) = UCase("wIrRmAddWt") Or _
              UCase(.ActiveControl.IdName) = UCase("wIrRmGrWt") Or _
              UCase(.ActiveControl.IdName) = UCase("wLsGrsWt") Then
                .ActiveControl = wInWt
            End If
        Case Is = UCase("GrdTxndB")
            If .ActiveControl.Col = 10 Then
                  
                  If .ActiveControl.Value(.ActiveControl.Row, "TdPrtKey") <> ctCurrPrtn Then DispMsg "Cannot Edit Locked/Previous Partition Entry", etError: Exit Sub
                  'Or mf_DtLocked = True
                  
                  If UCase$(.ADC("TLockYN")) = UCase$("Y") Then DispMsg "Cannot Add/Edit/Delete Locked Record", etError: Exit Sub
                  'If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
                  
                  If .ADC("TTC") = "BS" And .ADC("TtrfQty") <> 0# Then Exit Sub
                  
                    If ADC.Connection.RecSeek("Select TdTc From Txnd Where " + _
                               "TdCoCd='" + .ADC("TCoCd") + "' and TdTc='" + .ADC("TTc") + "' and " + _
                               "TdYy='" + .ADC("TYy") + "' and TdChr='" + .ADC("TChr") + "' and " + _
                               "TdNo=" + CStr(.ADC("TNo")) + " and " + _
                               "TdSr=" + CStr(.GrdTxndA.Value(.GrdTxndA.Row, "TdSr")) + " and " + _
                               "TdSrNo=" + CStr(.ActiveControl.Value(.ActiveControl.Row, "TdSrNo")) + " and TdPrtKey='" + ctCurrPrtn + "'") And _
                       (Not ADC.Connection.RecSeek("Select BYy from Bag where BCoCd='" + .ADC("TCoCd") + "' and " + _
                                    "BYy='" + .GrdTxndA.Value(.GrdTxndA.Row, "TdBYy") + "' and " + _
                                    "BChr='" + .GrdTxndA.Value(.GrdTxndA.Row, "TdBChr") + "' and " + _
                                    "BNo= " + CStr(.GrdTxndA.Value(.GrdTxndA.Row, "TdBNo")) + " and " + _
                                    "BLstYy='" + .ADC("TYy") + "' and " + _
                                    "BLstKey=" + CStr(.GrdTxndA.Value(.GrdTxndA.Row, "TdKey")) + " and BPrtKey='" + ctCurrPrtn + "'")) Then
                      DispMsg "Cannot Edit Since Not The Last Transaction Of The Bag", etError: Exit Sub
                    End If
                    
                    Dim wBFrLoc As String
                    If (.ADC("TTC") = "DT") = True Then
                      If gs_ValidPrdLocSeq <> "" Then
                        wBFrLoc = ADC.Connection.GetFldVal("Select BLoc from Bag where " + _
                                                 "BCoCd='" + .ADC("TCoCd") + "' and " + _
                                                 "BYy='" + .GrdTxndA.Value(.GrdTxndA.Row, "TdBYy") + "' and " + _
                                                 "BChr='" + .GrdTxndA.Value(.GrdTxndA.Row, "TdBChr") + "' and " + _
                                                 "BNo= " + CStr(.GrdTxndA.Value(.GrdTxndA.Row, "TdBNo")) + " and BPrtKey='" + ctCurrPrtn + "'")
                        If (InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(wBFrLoc) + ",") <= 0) = True Then
                          DispMsg "Cannot Edit record As Bag Loc. is Not Accessible to the User", etError: Exit Sub
                        End If
                      End If
                    End If
            
                    If .ActiveControl.Value(.ActiveControl.Row, "TdFbRmWt") <> "" Then
                      DispMsg "Rm Wt For Flute Bags Cannot Be Changed Directly. Click On 'Flute Bag Qty/Wt' To Change Wt.", etError
                      Exit Sub
                    End If
            
                .ActiveControl.Col = .ActiveControl.Col - 1
                .ActiveControl.Value(.ActiveControl.Row, "TdRmWt") = wInWt
                .ActiveControl.Col = .ActiveControl.Col + 1
            End If
        End Select
    
    
    
    Case Is = UCase("EmrFrmTxnRm")
        Select Case UCase(.ActiveControl.Name)
        Case Is = UCase("GrdTxndA")
            If .ActiveControl.Col = 11 Then
                If (.GrdTxndA.Value(.GrdTxndA.Row, "TdPrtKey") <> ctCurrPrtn) Then DispMsg "Cannot Edit Locked/Previous Partition Entry", etError: Exit Sub
                'Or mf_DtLocked = True
                If UCase$(.ADC("TLockYN")) = UCase$("Y") Then DispMsg "Cannot Add/Edit/Delete Locked Record", etError: Exit Sub
                'If mf_ValidLocSeq = False Then DispMsg "The User Does Not Have Access To One Of The Locs", etError: Exit Sub
            
                .ActiveControl.Col = .ActiveControl.Col - 1
                .ActiveControl.Value(.ActiveControl.Row, "TdRmWt") = wInWt
                .ActiveControl.Col = .ActiveControl.Col + 1
            End If
        End Select
    
    Case Is = UCase("EmrFrmTree")
        Select Case UCase(.ActiveControl.Name)
        Case Is = UCase("ATXT")
            If UCase(.ActiveControl.IdName) = UCase("TrFrMetWt") Or _
              UCase(.ActiveControl.IdName) = UCase("TrUsdMetWt") Or _
              UCase(.ActiveControl.IdName) = UCase("TrRecdMetWt") Or _
              UCase(.ActiveControl.IdName) = UCase("TrGrsPcWt") Then
                If .ADC("TrCstdt") = MWLib.EmptyDate Then DispMsg "Cannot Edit Value,when casting date is blank", etError: Exit Sub
                .ActiveControl = wInWt

              
            End If
        End Select
    End Select
    End With
  End If
  Exit Sub
errhld:
    DispMsg Err.Description, etError

End Sub

Private Function GetNumVal(strIn As String) As String
  Dim i As Integer
    For i = 1 To Len(strIn)
        If IsNumeric(Mid(strIn, i, 1)) Or Mid(strIn, i, 1) = "." Then GetNumVal = GetNumVal + Mid(strIn, i, 1)
    Next
End Function


Private Sub MultiPrcQt_Click()
  Call FormSelect("MultiPrcQT")
End Sub

Private Sub RepBvRm_Click()
  ' **** Zubin 2.11 **** '
  Call FormSelect("RepBvRm")
' **** Zubin 2.11 **** '
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

Private Sub RepLclInv_Click()
  Call FormSelect("RepLclInv")
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

Private Sub RepMultiPrcQt_Click()
  Call FormSelect("RepMultiPrcQt")
End Sub

Private Sub REPNETSTK_Click()
  Call FormSelect("RepNetStk")
End Sub
Private Sub RepRmReqSet_Click()
  Call FormSelect("RepRmReqSet")
End Sub
'MW.107 - New menu "Shipment Analysis" created
Private Sub RepShpAna_Click()
  Call FormSelect("RepShpAna")
End Sub

Private Sub RepPckLstInvBag_Click()
  Call FormSelect("RepPckLstInvBag")
End Sub

Private Sub RepTagPrnBagFg_Click()
  Call FormSelect("RepTagPrnBagFg")
End Sub
Private Sub RepTagPrnFg_Click()
  Call FormSelect("RepTagPrnFg")
End Sub






' ***** Manali 3.9.1 - 21/02/12 - RfId Module
Private Sub RfIdMst_Click()
  Call FormSelect("RfIdMst")
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

'indigo.164 new menu for Sale order sketch
Private Sub SalOrdSk_Click()
  Call FormSelect("SalOrdSk")
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
Private Sub Timer1_Timer()
'If StatusBar2.Panels("DiskSpace").text <> "" Then
'    StatusBar2.Panels("DiskSpace").text = ""
'Else
'    StatusBar2.Panels("DiskSpace").text = " LOW DISK SPACE ON LOG DRIVE. PERFORMANCE MAY SUFFER. CONTACT ADMINISTRATOR."
'End If
End Sub
Private Sub Tree_Click()
  Call FormSelect("Tree")
End Sub

Private Sub TXT_UCD_LostFocus()
    '3.11.4
    Dim wSrvrNm As String, i As Integer
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
    End If
    
    wSrvrNm = mo_MstConn.GetFldVal("Select MuSrvrNm From MUsrMst Where MuSysId='" + ctProjectName + "' And MuUsrId = '" + TXT_UCD + "'")
    
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
Private Sub Quot_Click()
  Call FormSelect("Quot")
End Sub
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
Private Sub PrdPtsSketch_Click()
  Call FormSelect("PrdPtsSketch")
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
Private Sub MA_Click()
  Call FormSelect("MA")
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
  ' **** Manali 3.8.0 - RepMetLs changed to RepMetLsWip
'  If WithinRTCRange = True Then
'    Call FormSelect("RepMetLs")
'  End If
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

Private Sub RepMetLsWip_Click()
  ' **** Manali 3.8.0 - RepMetLs changed to RepMetLsWip
  If WithinRTCRange = True Then
    Call FormSelect("RepMetLsWip")
  End If
End Sub


Private Sub RepMetLsFg_Click()
  If WithinRTCRange = True Then
    Call FormSelect("RepMetLsFg")
  End If
End Sub

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
' ** 3.9.0
Private Sub RepShpBillFile_Click()
  Call FormSelect("RepShpBillFile")
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
Private Sub RepValAddnDta_Click()
  Call FormSelect("RepValAddnDta")
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


Private Function GetVersion(ByRef ReturnMsg As String, ByRef databasever As String) As Boolean
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
    databasever = wDatVer
    ReturnMsg = wDatVer
  Else
    databasever = wDatVer
    ReturnMsg = "Cannot proceed. Exe Version is " & wExeVer & Chr(13) & "Database Version is " & wDatVer
    GetVersion = False
  End If
End Function
Private Function CheckModIntegrity() As String
  Dim wo_rsChkInt As MwfLib.MDORowSet, wiCtr As Integer, ws_ModulesFldStr As String

  CheckModIntegrity = ""

  Call SetModuleSeq
  
  ws_ModulesFldStr = ""
  Do While wiCtr <= UBound(gs_ModulesStr)
    ws_ModulesFldStr = ws_ModulesFldStr + "+" + gs_ModulesStr(wiCtr)
    wiCtr = wiCtr + 1
  Loop

  'Modules Check
  Set wo_rsChkInt = ADC.Connection.OpenResultset(" Select hChkIntgty, (HCoCd+':'+HName+':'" + ws_ModulesFldStr + ") as qModStatus " + _
                                               "  From Head " + _
                                               " Where HCd='" + ctSelfCmCd + "' And HCoCd='" + Trim(TxtCoCd.Text) + "'")
  With wo_rsChkInt
  If Not (.BOF Or .EOF) Then _
        If UCase(Decrypt(!hChkIntgty, GenCryptNew)) <> UCase(!qModStatus) Then _
        CheckModIntegrity = "Modules Check Failed in the Current Database, Contact Microway!!"
  End With
End Function

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
    If Not Ctl.SetLogInDets(TXT_UCD, ctProjectName, TxtCoCd, cmbSrvrNm.Text, cmbDbNm.Text) Then GoTo MDIEnd

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
    gs_PurgeDt = GetPurgeDt(TxtCoCd.Text, "")
                      
    Dim wErrMsg As String
    If CoChk(wErrMsg) = False Then
      MsgBox wErrMsg
      GoTo MDIEnd
    End If
    
    gs_HostNm = UCase$(GetHostNm)
    
    Dim wRetMsg As String, wSucc As Boolean
    'TEMPORARILY COMMENTED FOR TESTING OF 5.7.
'    If UCase(TXT_UCD) <> "MIC" Then
 '    wSucc = CheckSecurity(wRetMsg)
  '    If Not wSucc Then
   '     MsgBox wRetMsg
    '    GoTo MDIEnd
    '  End If
   ' End If
    

    Dim wDatabaseVer As String
    Dim wAutoUpdateDataReplaceErrMsg As String
    wSucc = GetVersion(wRetMsg, wDatabaseVer)
    gs_Ver = ""
    If wSucc Then
      gs_Ver = wRetMsg
    Else
      If UCase(TXT_UCD.Text) = "MIC" And _
          (wDatabaseVer = "5.0.0" Or wDatabaseVer = "5.0.1" Or wDatabaseVer = "5.0.2" Or wDatabaseVer = "5.1.0" Or wDatabaseVer = "5.2.0" Or wDatabaseVer = "5.2.1" Or wDatabaseVer = "5.2.2" Or wDatabaseVer = "5.2.3" Or wDatabaseVer = "5.2.4" Or wDatabaseVer = "5.2.5" Or wDatabaseVer = "5.2.6" Or wDatabaseVer = "5.2.7" Or wDatabaseVer = "5.2.8" Or wDatabaseVer = "5.2.9" Or wDatabaseVer = "5.2.10" Or wDatabaseVer = "5.3.0" Or wDatabaseVer = "5.4.0" Or wDatabaseVer = "5.4.1" Or wDatabaseVer = "5.4.2" Or wDatabaseVer = "5.4.3" Or wDatabaseVer = "5.4.4" Or wDatabaseVer = "5.5.0" Or wDatabaseVer = "5.5.1" Or wDatabaseVer = "5.5.2" Or wDatabaseVer = "5.5.3" Or wDatabaseVer = "5.6.0" Or wDatabaseVer = "5.6.1" Or wDatabaseVer = "5.6.2" Or wDatabaseVer = "5.6.3" Or wDatabaseVer = "5.7.0" Or wDatabaseVer = "5.7.1" Or wDatabaseVer = "5.7.2" Or wDatabaseVer = "5.7.3" Or wDatabaseVer = "5.7.4" Or wDatabaseVer = "5.7.5" Or wDatabaseVer = "5.7.6" Or _
          wDatabaseVer = "5.7.6" Or wDatabaseVer = "5.7.7" Or wDatabaseVer = "5.7.8" Or wDatabaseVer = "5.7.9" Or wDatabaseVer = "5.7.10" Or wDatabaseVer = "5.7.11" Or wDatabaseVer = "5.7.12" Or wDatabaseVer = "5.7.13" Or wDatabaseVer = "5.7.14" Or wDatabaseVer = "5.7.15" Or wDatabaseVer = "5.7.16" Or wDatabaseVer = "5.7.17" Or wDatabaseVer = "5.7.18" Or wDatabaseVer = "5.8.0") Then
          wSqlStrg = "SELECT * FROM PARAM WHERE PTYP= 'USR' AND PMCD ='" + TXT_UCD.Text + "' And PValidYn = 'Y'"
          Set wo_rsUsr = ADC.Connection.OpenRes(wSqlStrg)
          If wo_rsUsr.RecCount = 0 Then
            MsgBox ("Invalid User code")
            TXT_UCD.SetFocus
            GoTo MDIEnd
          End If
          If wo_rsUsr.FldValue("PValue") <> TXT_PWD.Text Then
            MsgBox ("Invalid Password. Cannot proceed.")
            TXT_PWD.SetFocus
            GoTo MDIEnd
          End If
             
             Dim ans As VbMsgBoxResult
             ans = MsgBox("Your database is currently at version " + wDatabaseVer + ". Would you like to upgrade " + "to latest version?", vbYesNo, "Upgrade?")
              If (ans = vbYes) Then
                ' CheckAndUpdateVersion will check if user has proper rights and update DB if chosen -
                  If Not CheckAndUpdateVersion(wRetMsg, wDatabaseVer) Then
                      MsgBox "Error trying to update your version. " + wRetMsg
                      GoTo MDIEnd
                  Else
                      MsgBox "Your Update is Complete"
                  End If
              Else
                MsgBox wRetMsg
                GoTo MDIEnd
              End If
        Else
          MsgBox wRetMsg
          GoTo MDIEnd
        End If
    End If
    
    If MWLib.IsItLaptop = False And ((MWLib.NoOfUsers) < ADC.Connection.GetFldVal("Select Count('x') From Param Where PTyp='USR' And PValidYN='Y' And PMCd<>'MIC' and PMCD<>'JMIC'")) Then
        MsgBox "Database Users are more than the Users Specified In the Ini File"
        GoTo MDIEnd
    End If

    wChkIntegrity = CheckModIntegrity()
    If Not wChkIntegrity = "" Then
      MsgBox wChkIntegrity
      GoTo MDIEnd
    End If
    Call HideAllMnus
    wSqlStrg = "SELECT * FROM PARAM WHERE PTYP= 'USR' AND PMCD ='" + TXT_UCD.Text + "' And PValidYn = 'Y'"
    Dim wl_UsrCnt As Long
    Set wo_rsUsr = ADC.Connection.OpenRes(wSqlStrg)
    If wo_rsUsr.RecCount = 0 Then
            MsgBox ("Invalid User code")
            TXT_UCD.SetFocus
            GoTo MDIEnd
    End If
    If wo_rsUsr.FldValue("PValue") <> TXT_PWD.Text Then
            MsgBox ("Invalid Password. Cannot log you in into the system.")
            TXT_PWD.SetFocus
            GoTo MDIEnd
    End If
     
    Dim wo_rsUL As MwfLib.MDORowSet
    
    Dim wi_delusr As Integer
    Dim ws_UsrLoggedInHostComputer As String
'6.1-49 UlSysCd='EMR' added
    ws_UsrLoggedInHostComputer = ADC.Connection.GetFldVal("Select UlHostNm From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlHostNm <> '" + gs_HostNm + "' and UlSysCd='EMR'")
'6.1-49 UlSysCd='EMR' added in select and delete query
    If ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlHostNm='" + gs_HostNm + "' and UlSysCd='EMR'") And Not ADC.Connection.RecSeek("Select 'x' From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlHostNm<>'" + gs_HostNm + "'") Then
       wi_delusr = MsgBox("User '" & UCase$(TXT_UCD.Text) & "' Already Logged In On This Machine. Delete User ?", vbYesNo, "User Check")
       If wi_delusr = vbYes Then
         ADC.Connection.Execute ("Delete From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' /* and UlCoCd = '" + UCase$(Trim$(TxtCoCd.Text)) + "' */ And UlHostNm='" + gs_HostNm + "' and UlSysCd='EMR'")
         'GoTo MDIAddUsrRec
       Else
         GoTo MDIEnd
       End If
    '6.1-49 UlSysCd='EMR' added
    ElseIf ADC.Connection.RecSeek("Select UlHostNm From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlHostNm <> '" + gs_HostNm + "'") Then
      ' We are knowingly seeking the database again cos we dont want to change the IF line above and change a whole lot - for just - hmm.. laziness
      ' One additional seek at the time of login when user is already logged in hmmm... we r gonnna live with it
       Dim ws_strHostNm As String
       '6.1-49 UlSysCd='EMR' added
       ws_strHostNm = ADC.Connection.GetFldVal("Select UlHostNm From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlHostNm <> '" + gs_HostNm + "'")
          MsgBox "Cannot Login. User '" & UCase$(TXT_UCD.Text) & "' Already Logged In On Computer '" + ws_strHostNm + "', Company " + UCase$(Trim$(TxtCoCd.Text)), vbOKOnly, "User Check"
          GoTo MDIEnd
    End If
        
    wl_UsrCnt = CLng(ADC.Connection.GetFldVal("Select Count (Distinct convert(char,UlUsrCd)+convert(char,UlHostNm))  from UsrLogin"))
    If wl_UsrCnt >= MWLib.NoOfUsers Then
      MsgBox "No of Users Allowed (" & MWLib.NoOfUsers & ") are Already Logged On", vbOKOnly, "User Count Check"
      GoTo MDIEnd
    End If
    
MDIAddUsrRec: '6.1-49 UlSysCd='EMR' added
        Set wo_rsUL = ADC.Connection.OpenResultset("Select * from UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlSysCd='EMR'")
        If Not (wo_rsUL Is Nothing) Then
          With wo_rsUL
            .AddNew
            .FldValue("UlUsrCd") = UCase$(Trim$(TXT_UCD.Text))
            .FldValue("UlCoCd") = UCase$(Trim$(TxtCoCd.Text))
            .FldValue("UlHostNm") = gs_HostNm
            .FldValue("UlDt") = ADC.Connection.SrvrDate
            .FldValue("UlTime") = ADC.Connection.SrvrTime
            .FldValue("UlSysCd") = "EMR"    '6.1-49
            .Update
          End With
      End If
'    End If
    '6.1-49 UlSysCd='EMR' added
    gs_UlIdNo = ADC.Connection.GetFldVal("Select UlIdNo From UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' and UlCoCd = '" + UCase$(Trim$(TxtCoCd.Text)) + "' " + _
                                        " and UlHostNm='" + gs_HostNm + "' and UlSysCd='EMR'")
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
                               "where HCoCd ='" + UCase(TxtCoCd.Text) + "' and HCd= '" + ctSelfCmCd + "'")
            
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
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from UsrAccess where uausrcd='" + UCase$(TXT_UCD.Text) + "' And UaMnuForm = '" + ms_MnuForm + "'" + _
                      " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.Text) & ",%' End" + ws_Mod)
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
            Set wo_rsusracc = ADC.Connection.OpenRes("select * from usraccess where uausrcd='" + UCase$(TXT_UCD.Text) + "' And UaMnuForm ='LST'" + _
                              " And ',' + UaValidCoCd + ',' Like Case When (UaValidCoCd = '') Then ',' + UaValidCoCd + ',' Else '%," & UCase$(TxtCoCd.Text) & ",%' End")
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
            'MnuTlBar.Visible = True     '*** (Jenny Color)
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
            
            Me.Caption = IIF(wMultiComp = True, wSelfCocdNm, gs_CoNm) + "  (" + Me.Caption + IIF(MWLib.IsItLaptop, " Laptop Version", "") + ", Db Svr= " + MWLib.DbSrvrNm + ", Db= " + MWLib.DbNm + "(" + IIF(TxtPartition.Text = "F", "Full", "Curr") + ")" + ", Rpt Svr= " + MWLib.RepDbSrvrNm + ", Rpt Db= " + MWLib.RepDbNm + ", Ver= " + wDbVer + ", Login Comp= " + gs_CoCd + ", Usr= " + TXT_UCD.Text + "," + DtShortStr + ")"
  
            SendKeys "%"
            
            ADC.Connection.Execute "Set Transaction Isolation Level Read UnCommitted"
'            End If
            
            

            OpenPort   ' ***** Manali 3.8.0 - Metler - MsComm Control Added

 
    If ADC.Connection.GetFldVal("Select Count(*) Cnt From Head Where HCoCd <> 'ZZZ'") = 1 Then
      gb_SingleCoCd = True
    Else
      gb_SingleCoCd = False
    End If
'    g_DateSetting = ReadSettings
    Set mFrmTrackStatus = New EmrFrmTrackStatus
    gs_DateFormatAtLogin = MWLib.GetDateFormatForEntry
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
Private Function CheckIntegrity(ByRef ErrMsg, ByVal databasever As String) As Boolean
On Error GoTo Err_Hndlr
  Dim wRs_ErrorRows As MDORowSet
  Dim wRes As Boolean
  wRes = True

  If databasever = "5.7.9" Or databasever = "5.7.10" Or databasever = "5.7.11" Or databasever = "5.7.12" Or databasever = "5.7.13" Or databasever = "5.7.14" Then
    Dim wrsRes As MDORowSet
    'Dim i As Integer, wSht As Integer, CpySht As Integer, sFilePath As String, wErrMsg As String
    Dim sFilePath As String, wErrMsg As String, wOutputPath As String
    Dim iFileNo As Integer, wRecStr As String, wFileNm As String
    
    wOutputPath = ADC.Connection.GetFldVal("Select PDesc225 from Param Where PTyp='OUTPUT' and PMCd='OUTPUT' ")
    If wOutputPath = "" Then
      MsgBox "Set a path in the Output folder to proceed with version update"
      CheckIntegrity = False
      GoTo Err_Hndlr
    End If
    
        
    iFileNo = FreeFile
    sFilePath = wOutputPath + "\Emr_5715_IntegrityChecks"
    
    If Dir(sFilePath, vbDirectory) = "" Then MkDir sFilePath
    If Dir(sFilePath + "\") <> "" Then Kill sFilePath + "\*.*"
    

    Set wrsRes = ADC.Connection.OpenRes("Select BCoCd, BYy, BChr, BNo from Bag where BYy='23' and not exists (Select 'a' from OrdDsg where OdIdNo= BOdIdNo and OdCoCd=BCoCd and OdTc=BOdTc and OdYy=BOdYy and OdChr = BOdChr and OdNo=BOdNo and BodSr=OdSr)")
    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Bag_Ord_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "BOdIdNo and OdIdNo not match"
        Print #iFileNo, ""
        Print #iFileNo, "BCoCd,BYy,BChr,BNo"
        Do While Not (.EOF Or .BOF)
          wRecStr = !BCoCd + "," + !BYy + "," + !BChr + "," + CStr(!BNo)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If
    
'    '''''''''''''1
    Set wrsRes = ADC.Connection.OpenRes("select FdCoCd ,FdTc ,FdYy ,FdChr ,FdNo ,FdSr from Fgd where FdYy='23' and not exists (Select 'a' from Bag where BCoCd =FdCoCd and BYy =FdBYy and BChr=FdBChr and BNo =FdBNo " + _
                                        "and BOdDmCd =FdDmCd and BOdSfx =FdSfx and BOdDmSz =FdDmSz)")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Fg_Bag_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Fg DsgCd/Sfx/Sz not match with Bag"
        Print #iFileNo, ""
        Print #iFileNo, "FdCoCd ,FdTc ,FdYy ,FdChr ,FdNo ,FdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !FdCoCd + "," + !FdTc + "," + !FdYy + "," + !FdChr + "," + CStr(!FdNo) + "," + CStr(!FdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'   '''''''''''''2
    Set wrsRes = ADC.Connection.OpenRes("Select BCoCd, BYy, BChr, BNo from Bag where BYy='23' and BCls = 'Y' and Round(BGrwt,3)<>0 and " + _
                                        "not exists (Select 'a' from InvFgd where BIdNo = IfBIdNo)")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Bag_Inv_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Bag Closed, NO Data in Invoice Bag"
        Print #iFileNo, ""
        Print #iFileNo, "BCoCd,BYy,BChr,BNo"
        Do While Not (.EOF Or .BOF)
          wRecStr = !BCoCd + "," + !BYy + "," + !BChr + "," + CStr(!BNo)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'   '''''''''''''3
    Set wrsRes = ADC.Connection.OpenRes("Select TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr from Txnd where TdAck='' and TdRmCd<>'' ")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\TdAck_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Record Not Acknowledged"
        Print #iFileNo, ""
        Print #iFileNo, "TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !TdCoCd + "," + !TdTc + "," + !TdYy + "," + !TdChr + "," + CStr(!TdNo) + "," + CStr(!TdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If
    
' '''''''''''''4
    Set wrsRes = ADC.Connection.OpenRes("Select TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr from Txnd where TdYy='23' and TdToBLoc <>'' " + _
                                        "and not exists (Select 'a' from Loc where LocCoCd=TdCoCd and LocCd=TdToBLoc)")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\ToBLoc_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Bag ToLocation Not Exist in Location Table"
        Print #iFileNo, ""
        Print #iFileNo, "TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !TdCoCd + "," + !TdTc + "," + !TdYy + "," + !TdChr + "," + CStr(!TdNo) + "," + CStr(!TdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'''''''''''''''5
    Set wrsRes = ADC.Connection.OpenRes("Select IdCoCd ,IdTc ,IdYy ,IdChr ,IdNo ,IdSr from InvDsg where IdYy='23' and not exists (Select 'a' from InvFgd join Bag on BIdNo=IfBIdNo where " + _
                                        "IfCoCd=IdCoCd and IfTc=IdTc and IfYy=IdYy and IfChr=IdChr and IfNo=IdNo and IfSr=IdSr and " + _
                                        "IdCoCd = BCoCd and IdExpOdTc = BOdTc and IdExpOdYy=BOdYy and IdExpOdChr = BOdChr and IdExpOdNo=BOdNo and IdExpOdSr=BOdSr)")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\InvDsg_Ord_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Invoice Design and Order Not Match"
        Print #iFileNo, ""
        Print #iFileNo, "IdCoCd ,IdTc ,IdYy ,IdChr ,IdNo ,IdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !IdCoCd + "," + !IdTc + "," + !IdYy + "," + !IdChr + "," + CStr(!IdNo) + "," + CStr(!IdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'  '''''''''''''6
    Set wrsRes = ADC.Connection.OpenRes("Select OdCoCd ,OdTc ,OdYy ,OdChr ,OdNo ,OdSr  from OrdDsg where OdYy='23' and OdTc='SO' and OdCls='Y' and OdOrdQty<>OdExpQty ")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\OdCls_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Order Closed Mismatch in OrdQty and ExpQty"
        Print #iFileNo, ""
        Print #iFileNo, "OdCoCd ,OdTc ,OdYy ,OdChr ,OdNo ,OdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !OdCoCd + "," + !OdTc + "," + !OdYy + "," + !OdChr + "," + CStr(!OdNo) + "," + CStr(!OdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If
    
' '''''''''''''7
    Set wrsRes = ADC.Connection.OpenRes("Select OrCoCd ,OrTc ,OrYy ,OrChr ,OrNo ,OrSr ,OrSrNo from OrdRm " + _
                                        "where  OrYy='23' and not exists (Select 'a' from OrdDsg where OdCoCd=OrCoCd and OdTc=OrTc and OdYy =OrYy " + _
                                        "and OdChr = OrChr and OdNo = OrNo and OdSr = OrSr and OdIdNo = OrOdIdNo )")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\OrdRm_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Order Rm Exist NO Order Design"
        Print #iFileNo, ""
        Print #iFileNo, "OrCoCd ,OrTc ,OrYy ,OrChr ,OrNo ,OrSr ,OrSrNo"
        Do While Not (.EOF Or .BOF)
          wRecStr = !OrCoCd + "," + !OrTc + "," + !OrYy + "," + !OrChr + "," + CStr(!OrNo) + "," + CStr(!OrSr) + "," + CStr(!OrSrNo)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

''''''''''''''8
    Set wrsRes = ADC.Connection.OpenRes("Select OlCoCd ,OlTc ,OlYy ,OlChr ,OlNo ,OlSr ,OlSrNo from OrdLab where OlYy='23' and  not exists (Select 'a' from OrdDsg " + _
                    "where OdCoCd=OlCoCd and OdTc=OlTc and OdYy =OlYy and OdChr = OlChr and OdNo = OlNo and OdSr = OlSr and OdIdNo = OlOdIdNo )")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\OrdLab_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Order Lab Exist NO Order Design"
        Print #iFileNo, ""
        Print #iFileNo, "OlCoCd ,OlTc ,OlYy ,OlChr ,OlNo ,OlSr ,OlSrNo"
        Do While Not (.EOF Or .BOF)
          wRecStr = !OlCoCd + "," + !OlTc + "," + !OlYy + "," + !OlChr + "," + CStr(!OlNo) + "," + CStr(!OlSr) + "," + CStr(!OlSrNo)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'   '''''''''''''10 exist
    Set wrsRes = ADC.Connection.OpenRes("Select TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr from Txnd where  TdYy='23' and not exists (Select 'a' from Txn where TdTIdNo=TIdNo)")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Txnd_Txn_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Txnd Exist NO Txn Records"
        Print #iFileNo, ""
        Print #iFileNo, "TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !TdCoCd + "," + !TdTc + "," + !TdYy + "," + !TdChr + "," + CStr(!TdNo) + "," + CStr(!TdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

'     '''''''''''''11 exist
    Set wrsRes = ADC.Connection.OpenRes("Select TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr from Txnd where TdYy='23' and  TdSrNo > 0 and not exists (Select 'a' from Txnd ForChk where ForChk.TdSrNo = 0 " + _
                  "and ForChk.TdCoCd = Txnd.TdCoCd and ForChk.TdTc = Txnd.TdTc and ForChk.TdYy=Txnd.TdYy and ForChk.TdChr = Txnd.TdChr " + _
                  "and ForChk.TdNo = Txnd.TdNo and ForChk.TdSr =Txnd.TdSr )")

    If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Txnd_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "TxndB exists NO TxndA"
        Print #iFileNo, ""
        Print #iFileNo, "TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !TdCoCd + "," + !TdTc + "," + !TdYy + "," + !TdChr + "," + CStr(!TdNo) + "," + CStr(!TdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If
    
'    '''''''''''''12 exist
   Set wrsRes = ADC.Connection.OpenRes("Select TdCoCd, TdTc, TdYy, TdChr, TdNo, TdSr from Txnd where TdYy='23' and TdBIdNo>0 and not exists (Select 'a' from Bag where TdBIdNo = BIdNo and TdCoCd = BCoCd and TdBYy=Byy and TdBChr = BChr and TdBNo=BNo) ")

   If wrsRes.RecCount > 0 Then
      wRecStr = ""
      wFileNm = sFilePath + "\Txnd_Bag_Mismatch.csv"
      Open wFileNm For Output As #iFileNo
      With wrsRes
        Print #iFileNo, "Txnd TdBIdNo and BIdNo not match"
        Print #iFileNo, ""
        Print #iFileNo, "TdCoCd, TdTc, TdYy, TdChr, TdNo,TdSr"
        Do While Not (.EOF Or .BOF)
          wRecStr = !TdCoCd + "," + !TdTc + "," + !TdYy + "," + !TdChr + "," + CStr(!TdNo) + "," + CStr(!TdSr)
          Print #iFileNo, wRecStr
          .MoveNext
        Loop
      End With
      Set wrsRes = Nothing
      wRes = False
      Close #iFileNo
    End If

  End If
  
'ChkDone:
  CheckIntegrity = wRes
  Exit Function
Err_Hndlr:
  If Err.Number = 70 Then
    MsgBox "Cannot Access file at path : " + sFilePath + ". Close All files If Already Open and Continue.", vbInformation
  Else
    wErrMsg = Err.Description
  End If
  Err.Clear
  
End Function
Private Function CheckAutoUpdateDataReplaceErr(ByRef ErrMsg, ByVal databasever As String) As Boolean
  Dim wRs_ErrorRows As MDORowSet
  Dim wRes As Boolean
  wRes = True

  If databasever = "5.0.0" Or databasever = "5.0.1" Or databasever = "5.0.2" Or databasever = "5.1.0" Or databasever = "5.2.0" Or databasever = "5.2.1" Or databasever = "5.2.2" Or databasever = "5.2.3" Or databasever = "5.2.4" Or databasever = "5.2.5" Or databasever = "5.2.6" Or databasever = "5.2.7" Or databasever = "5.2.8" Or databasever = "5.2.9" Or databasever = "5.2.10" Then
    wRes = Not ADC.Connection.RecSeek("Select 'a' from Param where PTyp ='LOCTYP' and PMCd ='R' and PSCd = ''")
    ErrMsg = ErrMsg + "Update to version 5.3.0 seem to have already run unsuccessfully on this database, please restore from backup to proceed"
     
    If ErrMsg = "" Then
      Set wRs_ErrorRows = ADC.Connection.OpenRes("Select * from Loc MainGuy where LocTyp = 'P' " + _
                                          "and exists " + _
                                        "(Select 'a' from loc where loctyp ='P' and loccocd = mainguy.LocCoCd and LocCd = 'R'+MainGuy.Loccd)")
      ErrMsg = ErrMsg + "Existing Location rows prevent creation of unique Process rows by adding 'R' to existing Codes. Location Codes:"
      Do While Not (wRs_ErrorRows.EOF Or wRs_ErrorRows.BOF)
        wRes = False
        ErrMsg = ErrMsg + "," + wRs_ErrorRows!LocCoCd + "/" + wRs_ErrorRows!LocCd
        wRs_ErrorRows.MoveNext
      Loop
    End If
  End If

  If ErrMsg <> "" Then GoTo ChkDone
  Dim rsWrongRmSzSetting As MDORowSet
  Set rsWrongRmSzSetting = ADC.Connection.OpenRes("Select PMCd, PSCd from Param where PTyp = 'RMSCtg' and PMCd in ('D','C') and exists " + _
        " (Select 'a' from Txnd where TdRmCd<>'' and PMCd = TdRmCd and PDesc225='Y' and PNum1<>1 and ((PNum1 <3 and TdRmSz3 <>0) or (PNum1<2 and (TdRmSz2<>0 or TdRmSz3<>0)))) ")
  Dim wSzStr As String
  If rsWrongRmSzSetting.RecCount > 0 Then
    wRes = False
    Do While Not (rsWrongRmSzSetting.EOF Or rsWrongRmSzSetting.BOF)
      wSzStr = wSzStr + "RmCtg/RMSCtg:" + rsWrongRmSzSetting!PMCd + "/" + rsWrongRmSzSetting!PSCd + vbCrLf
      ErrMsg = "Some of the Rm Sub Categories have no of sizes less than what is found in Transaction records. (wSzStr). Please contact Microway."
    Loop
  End If
  
  
ChkDone:
  CheckAutoUpdateDataReplaceErr = wRes
End Function
Private Function CheckSecurity(ByRef ErrMsg As String) As Boolean
    Dim wsql As String
    ErrMsg = ""
    wsql = "SELECT 'a' From master.sys.server_principals where upper(name) = upper('MwEmrRep')"
    If ADC.Connection.RecSeek(wsql) Then
      ErrMsg = "Critical System errors encountered, contact Microway"
      CheckSecurity = False
      Exit Function
    End If
    If ErrMsg = "" Then
      wsql = "SELECT sid,name FROM sys.server_principals " + _
              " WHERE is_disabled = 0 and NOT TYPE IN ( 'R')  " + _
               " AND NOT name IN ('##MS_PolicyEventProcessingLogin##', '##MS_PolicyTsqlExecutionLogin##') " + _
               " and name <> 'sa' and upper(name) <> upper('MW') and upper(name) <> upper('UsrRead') and upper(name) <> upper('MwEmrBackup')"
      If ADC.Connection.RecSeek(wsql) Then
        ErrMsg = "Critical System errors encountered, contact Microway"
        CheckSecurity = False
        Exit Function
      End If
    End If
    CheckSecurity = True
End Function
Private Function CheckAndUpdateVersion(ByRef ErrMsg, databasever As String) As Boolean
'  Check current version and can u upgrade
    Dim strCSV As String
    On Error GoTo ErrH
    Dim ans As VbMsgBoxResult
    
'    If DatabaseVer <> "5.0.0" Then
 '       ErrMsg = "Your database version is " + DatabaseVer + ". You cannot update directly through this option.Contact Microway"
  '      CheckAndUpdateVersion = False
   '     Exit Function
'    End If
    
    Dim ActiveConnStr As String

    Dim rsActiveConn As MDORowSet
    Set rsActiveConn = ADC.Connection.OpenRes _
    ("select hostname from sys.sysprocesses where dbid =  db_id() and ((upper(hostname) ='" + UCase(gs_HostNm) + "' and upper(loginame) <> 'MW') or (upper(hostname) <> '" + UCase(gs_HostNm) + "'))")
    Do While Not (rsActiveConn.EOF Or rsActiveConn.BOF)
      ActiveConnStr = ActiveConnStr + Trim(rsActiveConn!HostName) + ";"
      rsActiveConn.MoveNext
    Loop
    
    If rsActiveConn.RecCount > 0 Then
        ErrMsg = "There are active connections to database " + MWLib.DbNm + " from " + ActiveConnStr + ". Please terminate all the connections to be able to Upgrade"
        CheckAndUpdateVersion = False
        Exit Function
    End If
    
        
'    If adc.Connection.RecSeek("Select 'a' from sys.databases where (is_published = 1 or is_subscribed = 1 or is_merge_published = 1 or is_distributor = 1) and Name = db_name()") Then
 '       ErrMsg = "Replication seems to be running on this database. Please stop replication and try again to Autoupdate"
  '      CheckAndUpdateVersion = False
   '     Exit Function
'    End If
    
    If databasever <> "5.0.0" And databasever <> "5.0.1" And databasever <> "5.0.2" And databasever <> "5.1.0" And databasever <> "5.2.0" And databasever <> "5.2.1" And databasever <> "5.2.2" And databasever <> "5.2.3" And databasever <> "5.2.4" And databasever <> "5.2.5" And databasever <> "5.2.6" And databasever <> "5.2.7" And databasever <> "5.2.8" And databasever <> "5.2.9" And databasever <> "5.2.10" And databasever <> "5.3.0" And databasever <> "5.4.0" And databasever <> "5.4.1" And databasever <> "5.4.2" And databasever <> "5.4.3" And databasever <> "5.4.4" And databasever <> "5.5.0" And databasever <> "5.5.1" And databasever <> "5.5.2" And databasever <> "5.5.3" And databasever <> "5.6.0" And databasever <> "5.6.1" And databasever <> "5.6.2" And databasever <> "5.6.3" And databasever <> "5.7.0" And databasever <> "5.7.1" And databasever <> "5.7.2" And databasever <> "5.7.3" And databasever <> "5.7.4" And databasever <> "5.7.5" And databasever <> "5.7.6" _
          And databasever <> "5.7.7" And databasever <> "5.7.8" And databasever <> "5.7.9" And databasever <> "5.7.10" And databasever <> "5.7.11" And databasever <> "5.7.12" And databasever <> "5.7.13" And databasever <> "5.7.14" And databasever <> "5.7.15" And databasever <> "5.7.16" And databasever <> "5.7.17" And databasever <> "5.7.18" And databasever <> "5.8.0" Then
        ErrMsg = "Your database version is " + databasever + ". You cannot Upgrade version directly through this option.Contact Microway"
        CheckAndUpdateVersion = False
        Exit Function
    End If
    Dim wErrMsg As String
    If Not CheckAutoUpdateDataReplaceErr(wErrMsg, databasever) Then
        ErrMsg = "Existing Data violates Data rule check for Autoupdate. " + wErrMsg
        CheckAndUpdateVersion = False
        Exit Function
    End If

    If Not CheckIntegrity(wErrMsg, databasever) Then
        ErrMsg = "Existing Data violates Data rule check for Autoupdate. " + wErrMsg
        CheckAndUpdateVersion = False
        Exit Function
    End If
    

    Dim Upd_510_520_PrimaryKeyChange As String
    Upd_510_520_PrimaryKeyChange = "'SYYMM','DMYYMM'"
    
    Dim Upd_501_502_PrimaryKeyChange As String
    Upd_501_502_PrimaryKeyChange = "'INVFGD','DMYYMM'"
    
    Dim Upd_500_501_PrimaryKeyChange As String
    Upd_500_501_PrimaryKeyChange = ""

    If databasever = "5.0.0" Then
      If ADC.Connection.RecSeek("Select 'a' from sys.databases where (is_published = 1 or is_subscribed = 1 or is_merge_published = 1 or is_distributor = 1) and Name = db_name()") Then
      
       If ADC.Connection.RecSeek("Select 'a' from sysarticles where upper(name) in (" + Upd_500_501_PrimaryKeyChange + Upd_501_502_PrimaryKeyChange + ")") Then
        ErrMsg = "Tables " + Upd_500_501_PrimaryKeyChange + Upd_501_502_PrimaryKeyChange + " have Primary Key change and some of these tables are published for Replication. Please remove these from Replication and try again"
        CheckAndUpdateVersion = False
        Exit Function
       End If
      End If
     

      strCSV = StrConv(LoadResData("Upd_500_501", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_501_502", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_502_510", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_510_520", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_520_521", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
     If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
        If databasever = "5.0.1" Then
      If ADC.Connection.RecSeek("Select 'a' from sys.databases where (is_published = 1 or is_subscribed = 1 or is_merge_published = 1 or is_distributor = 1) and Name = db_name()") Then
       If ADC.Connection.RecSeek("Select 'a' from sysarticles where upper(name) in (" + Upd_501_502_PrimaryKeyChange + ")") Then
        ErrMsg = "Tables " + Upd_501_502_PrimaryKeyChange + " have Primary Key change and some of these tables are published for Replication. Please remove these from Replication and try again"
        CheckAndUpdateVersion = False
        Exit Function
       End If
      End If
      strCSV = StrConv(LoadResData("Upd_501_502", "UPDATESCRIPTS"), vbUnicode)

      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_502_510", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_510_520", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_520_521", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    
    If databasever = "5.0.2" Then
      strCSV = StrConv(LoadResData("Upd_502_510", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
    
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_510_520", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_520_521", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    
    If databasever = "5.1.0" Then
      strCSV = StrConv(LoadResData("Upd_510_520", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_520_521", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    
    If databasever = "5.2.0" Then
      strCSV = StrConv(LoadResData("Upd_520_521", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    
    If databasever = "5.2.1" Then
      strCSV = StrConv(LoadResData("Upd_521_522", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
         strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
         ErrMsg = RunScript(strCSV)
       End If
       If ErrMsg = "" Then
         strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
         ErrMsg = RunScript(strCSV)
       End If
       If ErrMsg = "" Then
         strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
         ErrMsg = RunScript(strCSV)
       End If
       If ErrMsg = "" Then
         strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
         ErrMsg = RunScript(strCSV)
       End If
       If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
       If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.2" Then
        strCSV = StrConv(LoadResData("Upd_522_523", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.3" Then
      strCSV = StrConv(LoadResData("Upd_523_524", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.4" Then
        strCSV = StrConv(LoadResData("Upd_524_525", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.5" Then
      strCSV = StrConv(LoadResData("Upd_525_526", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
           If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.6" Then
        strCSV = StrConv(LoadResData("Upd_526_527", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If

    Call Tp(ErrMsg, databasever)
    If (databasever = "5.6.3") Then
      strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.7.0") Then
      strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.7.1") Then
      strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    If (databasever = "5.7.2") Then
      strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.7.3") Then
      strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    If (databasever = "5.7.4") Then
      strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.7.5") Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
              If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.7.6") Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

    End If
    If (databasever = "5.7.7") Then
      strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    If (databasever = "5.7.8") Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
        
    End If
    
    If (databasever = "5.7.9") Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

    End If
    If (databasever = "5.7.10") Then
          strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
          
    End If
    If (databasever = "5.7.11") Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    
    If (databasever = "5.7.12") Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5713_580", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.13") Then
        strCSV = StrConv(LoadResData("Upd_5713_5714", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5714_5715", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5715_5716", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5716_5717", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5717_5718", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.14") Then
        strCSV = StrConv(LoadResData("Upd_5714_5715", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5715_5716", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5716_5717", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5717_5718", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.15") Then
        strCSV = StrConv(LoadResData("Upd_5715_5716", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5716_5717", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5717_5718", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.16") Then
        strCSV = StrConv(LoadResData("Upd_5716_5717", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5717_5718", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.17") Then
        strCSV = StrConv(LoadResData("Upd_5717_5718", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    End If
    If (databasever = "5.7.18") Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
    End If
    If (databasever = "5.8.0") Then
        strCSV = StrConv(LoadResData("Upd_5718_610", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
    End If
    
    Call SetModuleSeq
    If ErrMsg = "" Then
      Call UpdIntegrity(ADC.Connection)
      ADC.Connection.Execute ("Insert into TrgDisa values ('Param')")
      ADC.Connection.Execute ("Update Param set Pdesc225 = '6.1.0' where Ptyp = 'VER'")
      ADC.Connection.Execute ("Delete from TrgDisa")
      CheckAndUpdateVersion = True
      Exit Function
    End If
ErrH:
  ADC.Connection.Execute ("Delete from TrgDisa")
  ADC.Connection.Execute ("Insert into TrgDisa values ('Param')")
  ADC.Connection.Execute ("Update Param set Pdesc225 = '0.0.0' where Ptyp = 'VER'")
  ADC.Connection.Execute ("Delete from TrgDisa")
  ErrMsg = Err.Description + ":" + ErrMsg
  CheckAndUpdateVersion = False
End Function
Private Function RunScript(ByVal strCSV As String) As String
On Error GoTo ErrH
Dim ErrMsg As String
    Dim fullCommand As String
    Dim curCommand As String
    Dim i As Long
    Dim stopPos As Long
    stopPos = 0
    fullCommand = strCSV
    Do While (fullCommand <> "")
      i = InStr(1, UCase(fullCommand), "GOGO")
      stopPos = IIF(i > 0, i - 1, Len(fullCommand))
      curCommand = Mid(fullCommand, 1, stopPos)
      If InStr(1, curCommand, "@Cell") > 0 Then
        Dim W As String
        W = "abded"
      End If
      ADC.Connection.Execute curCommand
      fullCommand = IIF(i > 0, Mid(fullCommand, i + 3 + 2), "")
    Loop
    RunScript = ""
    Exit Function
ErrH:
  ErrMsg = Err.Description + "Command: " + curCommand
  RunScript = ErrMsg
End Function

  
  Public Sub Ctl_HandleError(ByVal pv_ErrMsg As String, ByVal pv_ErrorType As MwfCtl.en_ErrorType)

  If UCase(Trim(pv_ErrMsg)) = UCase(Trim("Out of Memory")) Then pv_ErrMsg = Trim(pv_ErrMsg) + ". Please Reduce The Scope."
  If g_InProcessingMode And (pv_ErrorType = etInfo Or pv_ErrorType = etWarning) Then
    'DONOTHING
  Else
    StatusBar2.Panels("ErrMsg").Text = pv_ErrMsg
  End If
  If pv_ErrorType = etInfo And g_InProcessingMode Then
   If IsItUpdateEmrProcessMsg(pv_ErrMsg) Then
    StatusBar2.Panels("ErrMsg").Text = pv_ErrMsg
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
Private Sub Tp(ByRef ErrMsg, databasever As String)
    Dim strCSV As String

    If databasever = "5.2.7" Then
        strCSV = StrConv(LoadResData("Upd_527_528", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
            strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
            ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
        
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.2.8" Then
        strCSV = StrConv(LoadResData("Upd_528_529", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        
        If ErrMsg = "" Then
            strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
            ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
        
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    If databasever = "5.2.9" Then
        strCSV = StrConv(LoadResData("Upd_529_5210", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        
        If ErrMsg = "" Then
            strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
            ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
    

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If databasever = "5.2.10" Then
        strCSV = StrConv(LoadResData("Upd_5210_530", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        
        If ErrMsg = "" Then
            strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
            ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If databasever = "5.3.0" Then
        strCSV = StrConv(LoadResData("Upd_530_540", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        If ErrMsg = "" Then
            strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
            ErrMsg = RunScript(strCSV)
        End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      
    End If
    'sri - Specifically changing this behavior as DB updates of 5.4.1 and 5.4.2 and 5.4.3 have been merged
    'into trigger changes of 5.5.0. All changes are there except one param data change done in 5.4.1 to
    'update Repair Character to N
    If (databasever = "5.4.0" Or databasever = "5.4.1" Or databasever = "5.4.2" Or databasever = "5.4.3") Then
        If ErrMsg = "" Then
          strCSV = StrConv(LoadResData("Upd_543_544", "UPDATESCRIPTS"), vbUnicode)
          ErrMsg = RunScript(strCSV)
        End If
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.4.4") Then
        strCSV = StrConv(LoadResData("Upd_544_550", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
          If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

    End If
    If (databasever = "5.5.0") Then
        strCSV = StrConv(LoadResData("Upd_550_551", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.5.1") Then
        strCSV = StrConv(LoadResData("Upd_551_552", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.5.2") Then
        strCSV = StrConv(LoadResData("Upd_552_553", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.5.3") Then
        strCSV = StrConv(LoadResData("Upd_553_560", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If (databasever = "5.6.0") Then
      strCSV = StrConv(LoadResData("Upd_560_561", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      

    End If
    If (databasever = "5.6.1") Then
      strCSV = StrConv(LoadResData("Upd_561_562", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
                  If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_575_576", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
            If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_576_577", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
    End If
    If (databasever = "5.6.2") Then
      strCSV = StrConv(LoadResData("Upd_562_563", "UPDATESCRIPTS"), vbUnicode)
      ErrMsg = RunScript(strCSV)
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_563_570", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If

      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_570_571", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_571_572", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_572_573", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_573_574", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_574_575", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_577_578", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_578_579", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_579_5710", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5710_5711", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5711_5712", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      If ErrMsg = "" Then
        strCSV = StrConv(LoadResData("Upd_5712_5713", "UPDATESCRIPTS"), vbUnicode)
        ErrMsg = RunScript(strCSV)
      End If
      
            
    End If

End Sub

Private Sub CmdCan_Click()
    Unload Me
End Sub
Private Sub csc_Click()
    EmrFrmMDI.Arrange vbCascade
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
     Call SetMDIFrm(Me)
    
    takecloseoff hwnd
    Set MWLib = New MWCLS_GEN

    TxtPartition.Visible = False
    
    ms_MnuForm = "MDI"

    wb_ApplyClr = False
    TlBar.Visible = False
    Picture1.Picture = Me.Picture
    
    SaveOriginalColors
  
    SetMDISysColorsBefLogin
      
    If Not MWLib.SetDateFormat Then GoTo MDIEnd
    
    MWLib.SetLoginParameters "Mw", "EmrSQLMw007*99", "", ""
    'Ctl.SetLoginParameters "UsrApp", "UsrAppPwd", "AppRole", "AppRolePwd"
    Ctl.SetLoginParameters "Mw", "EmrSQLMw007*99", "", ""
    
    Set mo_MstConn = MWLib.GetMwMstConnection(ctProjectName)
    
    If mo_MstConn Is Nothing Then
        DispMsg "Unable to Connect to Master Server", etError
        Exit Sub
     End If

     Call FillSrvrCmb
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
    SetMDIPic
  Else
     RestoreColors
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  RestoreColors
   Call SetRecClrSet(Nothing)
   Call SetMDIFrm(Nothing)
  Set mo_MDIPic = Nothing
  Set mo_CsTlBarPic = Nothing
  Set mo_TlBarColl = Nothing
   Set MWLib = Nothing
  Set HlpList = Nothing
  Set gColTgFldTyp = Nothing
  Set mwhlp = Nothing
  '*** (09/08/05)
  
 End Sub
Private Sub quit_Click()
    Unload mFrmTrackStatus
    Set mFrmTrackStatus = Nothing
    If Not (Me.ActiveForm Is Nothing) Then
       Call Ctl_HandleError("Close all Child forms before quitting", etError)
    Else
       Set Ctl.ActScrn = Nothing
       Set Ctl.MDIApp = Nothing
       
        Set MWLib = Nothing   '*** (Bef 24/11/06)
       '*** (Jen 04/02/08)

        Dim wo_rsUL As MwfLib.MDORowSet
        '*** Jay 3.1.0_1 [UlCoCd Added]
        '6.1-49 UlSysCd='EMR' added
        Set wo_rsUL = ADC.Connection.OpenResultset("Select * from UsrLogin Where UlUsrCd='" + UCase$(Trim$(TXT_UCD.Text)) + "' " + _
                            "and UlCoCd='" + UCase$(Trim$(TxtCoCd)) + "' And UlHostNm = '" + gs_HostNm + "' and UlSysCd='EMR'")
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
        ClosePort   ' ***** Manali 3.8.0 - Metler - MsComm Control Added
       Unload Me
    End If
End Sub
Private Sub til_Click()
    EmrFrmMDI.Arrange vbTileVertical
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
Dim TESTVAR As Variant
  On Error GoTo GenMnuErr
    Dim rs_UsrAcc As MwfLib.MDORowSet
    Dim wo_Ctl As Control
    MWLib.BeginProcess Me, "Menu Generation Is In Progress"
    'And UCase$(wo_Ctl.Name) <> UCase$("IdeDfn")
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
              UCase$(wo_Ctl.Name) <> "TIL" And UCase$(wo_Ctl.Name) <> "GENMNU" And UCase$(wo_Ctl.Name) <> "QUIT" And UCase$(wo_Ctl.Name) <> "TRACK" And UCase$(wo_Ctl.Name) <> "MNUTLBAR" _
              And UCase$(wo_Ctl.Name) <> UCase$("RepDeclnLetMulti") Then
' ###########################################  Manoj  ###########################################
                'Debug.Print wo_Ctl.Name
                rs_UsrAcc.AddNew
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
                'rs_UsrAcc!UaMetlrCd = ""    '***** Manali 3.8.0 - Metlr Cd
                TESTVAR = rs_UsrAcc!UaMnuCd
                rs_UsrAcc.Update
            End If
        End If
    Next
    
    ' **** Manali 3.8.0 - IsNull(UA1.UaTag, '') Replaced as IsNull(UA1.UaTag, UaTag)
    ADC.Connection.Execute ("Update UsrAccess Set UaTag = IsNull((Select isNull(UA1.UaTag , '') From UsrAccess UA1 " + _
        "Where UA1.UaMnuCd= UsrAccess.UaMnuCd And UA1.UaMnuForm = UsrAccess.UaMnuForm  And " + _
        "UA1.UAUSRCD= '" + ctGenMenuUsr + "'), UaTag) Where UAMNUFORM = '" + ms_MnuForm + "'")
        
    ADC.Connection.Execute ("Update UsrAccess set UaMnuDesc = " + _
                             "isnull((Select top 1 UaMnuDesc from UsrAccess UsrAccessForParent where UsrAccessForParent.UaMnuCd=ChildTable.UaMnuCd and UsrAccessForParent.UaMnuForm = ChildTable.uamnuform  and " + _
                             "UsrAccessForParent.UaUsrCd in ('MIC','JMIC')),ChildTable.UaMnuDesc) " + _
                             " From " + _
                            "  UsrAccess ChildTable " + _
                            "    Where " + _
                            "  UausrCd<>'MIC' and UausrCd<>'JMIC'")


    DispMsg "Menu on '" + ms_MnuForm + "' From Succesfully Generated for '" + ctGenMenuUsr + "'", etInfo
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
    gs_CoCd = UCase(EmrFrmMDI.TxtCoCd.Text)
    gs_CoNm = ADC.Connection.GetFldVal("Select HName from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
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
  StatusBar2.Panels("InMail").Text = IIF(wiMsgCnt <> 0, CStr(wiMsgCnt), "")
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
Public Sub DtFormatError()
  Unload Me
End Sub
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
  'If MailSignIn.Visible = True Then MailSignIn = Not gb_MailAct
  'If MailSignOut.Visible = True Then MailSignOut = gb_MailAct
  'If SndMail.Visible = True Then SndMail = gb_MailAct
  'StatusBar2.Panels("InMail").Visible = gb_MailAct
End Sub
' ###########################################  Manoj  ###########################################

' Procedure To Get Tag & Code of Menu
Private Sub PrnMnus()
    Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is Menu Then
           If Ctl.Caption <> "DUM" Then
               ' Debug.Print Ctl.Tag + vbTab + vbTab + vbTab + Ctl.Name
           End If
        End If
    Next
End Sub

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

'Private Sub DiskSpaceChk()
''3.11.2 - Disk Space Check on Log File Drive
'    Dim wDSpace As MDORowSet
'    Set wDSpace = adc.Connection.OpenRes(" Set DateFormat DMY; " + _
'                                         " Select 'x' from diskinfo Where DiLogDrvFreeSpace<DiIdxSize*3 And ModDt=(Select Max(ModDt) from diskinfo) And DateDiff(day, ModDt, GetDate())<10")
'    If wDSpace.RecCount > 0 Then _
'        StatusBar2.Panels("DiskSpace").text = " LOW DISK SPACE ON LOG DRIVE. PERFORMANCE MAY SUFFER. CONTACT ADMINISTRATOR."
'    Set wDSpace = Nothing
'
'    Set wDSpace = adc.Connection.OpenRes(" Set DateFormat DMY; " + _
'                                         " Select 'x' from diskinfo Where ModDt=(Select Max(ModDt) from diskinfo) And DateDiff(day, ModDt, GetDate())>10")
'    If wDSpace.RecCount > 0 Or Not adc.Connection.RecSeek("Select 'x' from diskinfo") Then _
'        StatusBar2.Panels("DiskSpace").text = " DISK STATUS NOT UPDATED. CONTACT ADMINISTRATOR."
'
'    Set wDSpace = Nothing
'End Sub

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
                                       "  and  MsSrvrNm ='" + cmbSrvrNm.Text + "'")

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
'VK.5-added menu for So Pending Rm detail
'mw.104-SO Rm pending report moved from AddOns menu  to MDi
Private Sub RepSOPndRM_Click()
  Call FormSelect("RepSOPndRM")
End Sub

'6.1-15
Private Sub PdcmUpd_Click()
  Call FormSelect("PdcmUpd")
End Sub
'6.1 hlp
Private Sub hlp_Click()
  Dim wPath As String
  On Error GoTo ErrH
  wPath = ADC.Connection.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\Help\" + gs_ActFrmMenucd + ".pdf"
  OLE1.CreateLink wPath
  OLE1.DoVerb
  Exit Sub
ErrH:
  If Err.Number = 31031 Then
    DispMsg "File/Folder (" + wPath + ") Not Found", etError
  Else
    DispMsg Err.Description, etError
  End If

End Sub

