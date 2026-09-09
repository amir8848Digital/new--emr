VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmCustMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11040
   ScaleMode       =   0  'User
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSC 
      Height          =   1530
      Left            =   2760
      TabIndex        =   217
      Top             =   7800
      Visible         =   0   'False
      Width           =   5385
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   88
         Left            =   1785
         TabIndex        =   221
         ToolTipText     =   "Enter To Design Code"
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WSCDMCDTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   89
         Left            =   885
         TabIndex        =   218
         ToolTipText     =   "Enter From Design Category"
         Top             =   690
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCDMCTGFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   78
         Left            =   880
         TabIndex        =   219
         ToolTipText     =   "Enter To Design Category"
         Top             =   975
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCDMCTGTO"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   91
         Left            =   1785
         TabIndex        =   220
         ToolTipText     =   "Enter From Design Code"
         Top             =   690
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WSCDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   298
         Left            =   4245
         TabIndex        =   222
         ToolTipText     =   "Enter From Row"
         Top             =   690
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "####0"
         MaxLength       =   5
         DataType        =   1
         IdName          =   "WSCIDSRFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   299
         Left            =   4245
         TabIndex        =   223
         ToolTipText     =   "Enter To Row"
         Top             =   975
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Alignment       =   2
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCIDSRTO"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   75
         Left            =   60
         TabIndex        =   229
         Top             =   690
         Width           =   585
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Row"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   228
         Left            =   4245
         TabIndex        =   228
         Top             =   420
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Scope Option For Selecting Records Into Cust Design Grid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   -30
         TabIndex        =   227
         Top             =   90
         Width           =   5130
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   74
         Left            =   60
         TabIndex        =   226
         Top             =   975
         Width           =   585
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Ctg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   73
         Left            =   880
         TabIndex        =   225
         Top             =   420
         Width           =   825
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Design"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   72
         Left            =   1830
         TabIndex        =   224
         Top             =   420
         Width           =   735
      End
   End
   Begin MwfCtl.MWCTL_BTN1 CmdSC 
      Height          =   495
      Left            =   5040
      TabIndex        =   216
      ToolTipText     =   "Option To Select The Scope Options For The Cust Design Entry"
      Top             =   9300
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Sco&pe"
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
   Begin MwfCtl.MWCTL_BTN1 cmdImpExc 
      Height          =   495
      Left            =   5790
      TabIndex        =   190
      ToolTipText     =   "Option To Copy from Excel"
      Top             =   9300
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Imp E&xcel"
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
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   88
      Top             =   9360
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRm 
      Height          =   495
      Left            =   6510
      TabIndex        =   156
      ToolTipText     =   "Option To Copy Alloy Details From Another Customer"
      Top             =   9300
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy &RM"
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpy 
      Height          =   495
      Left            =   7260
      TabIndex        =   120
      ToolTipText     =   "Option To Copy Alloy Details From Another Customer"
      Top             =   9300
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy Alloy"
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   8550
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   9315
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   -30
      TabIndex        =   94
      Top             =   30
      Width           =   3210
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   1
         Left            =   1110
         TabIndex        =   0
         ToolTipText     =   "Enter Customer Code"
         Top             =   -15
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         DataField       =   "CmCd"
         IdName          =   "CMCD"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   1
         Top             =   -15
         Visible         =   0   'False
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   503
         Enabled         =   0   'False
         MaxLength       =   2
         DataType        =   4
         DataField       =   "CmCtg"
         IdName          =   "CMCTG"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Cd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   95
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9915
      Left            =   -120
      TabIndex        =   93
      Top             =   0
      Width           =   15315
      Begin VB.Frame FraAddr 
         Height          =   1155
         Left            =   240
         TabIndex        =   207
         Top             =   8280
         Width           =   9690
         Begin MwfCtl.MWCTL_BTN1 CmdAddrOk 
            Height          =   525
            Left            =   8835
            TabIndex        =   211
            Top             =   390
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   926
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&OK"
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
         Begin MwfCtl.MWCTL_MED TXTADDR 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   208
            ToolTipText     =   "Enter Address1"
            Top             =   180
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "GrdCustStore(CSTADD1)"
            IdName          =   "WCSTADD1"
         End
         Begin MwfCtl.MWCTL_MED TXTADDR 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   209
            ToolTipText     =   "Enter Address2"
            Top             =   485
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "GrdCustStore(CSTADD2)"
            IdName          =   "WCSTADD2"
         End
         Begin MwfCtl.MWCTL_MED TXTADDR 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   210
            ToolTipText     =   "Enter Address3"
            Top             =   770
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   503
            MaxLength       =   45
            ReCalcParent    =   "GrdCustStore(CSTADD3)"
            IdName          =   "WCSTADD3"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Address1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   68
            Left            =   120
            TabIndex        =   214
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Address2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   69
            Left            =   120
            TabIndex        =   213
            Top             =   485
            Width           =   1065
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Address3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   70
            Left            =   120
            TabIndex        =   212
            Top             =   770
            Width           =   1065
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9510
         Left            =   120
         TabIndex        =   96
         Top             =   30
         Width           =   15150
         Begin TabDlg.SSTab TabCustDet 
            Height          =   8520
            Left            =   0
            TabIndex        =   4
            Top             =   630
            Width           =   15090
            _ExtentX        =   26617
            _ExtentY        =   15028
            _Version        =   393216
            Tabs            =   9
            Tab             =   3
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "<&1> Cust Details"
            TabPicture(0)   =   "EmrFrmCustMst.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraTabCustDet(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2> Instructions"
            TabPicture(1)   =   "EmrFrmCustMst.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FraTabCustDet(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Factors"
            TabPicture(2)   =   "EmrFrmCustMst.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraTabCustDet(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4> Cust Design"
            TabPicture(3)   =   "EmrFrmCustMst.frx":0054
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "FraTabCustDet(3)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5> Cust Alloy"
            TabPicture(4)   =   "EmrFrmCustMst.frx":0070
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "FraTabCustDet(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "<&6> Cust Rm"
            TabPicture(5)   =   "EmrFrmCustMst.frx":008C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "FraTabCustDet(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "<&7> Cust History"
            TabPicture(6)   =   "EmrFrmCustMst.frx":00A8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "FraTabCustDet(6)"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "<&8> Cust Prd Ins"
            TabPicture(7)   =   "EmrFrmCustMst.frx":00C4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "FraTabCustDet(7)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "<&9> Cust Store"
            TabPicture(8)   =   "EmrFrmCustMst.frx":00E0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "FraTabCustDet(8)"
            Tab(8).Control(0).Enabled=   0   'False
            Tab(8).ControlCount=   1
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   6915
               Index           =   8
               Left            =   -74940
               TabIndex        =   205
               Top             =   420
               Width           =   14880
               Begin MwfCtl.MWCTL_FLX GrdCustStore 
                  Height          =   6765
                  Left            =   90
                  TabIndex        =   206
                  Top             =   90
                  Width           =   14835
                  _ExtentX        =   26167
                  _ExtentY        =   11933
                  Cols            =   11
                  colname1        =   "CSTCD"
                  heading1        =   "Store Cd                           "
                  datafld1        =   "CstCd"
                  datatype1       =   4
                  maxlength1      =   12
                  tooltiptext1    =   "Enter Store Code"
                  colname2        =   "CSTNAME"
                  heading2        =   "Store Name                                                                                      "
                  datafld2        =   "CstName"
                  maxlength2      =   45
                  tooltiptext2    =   "Enter Store Name"
                  colname3        =   "CSTTEL"
                  heading3        =   "Tel                                            "
                  datafld3        =   "CstTel"
                  maxlength3      =   20
                  tooltiptext3    =   "Enter Telephone Number"
                  colname4        =   "CSTFAX"
                  heading4        =   "Fax                                            "
                  datafld4        =   "CstFax"
                  maxlength4      =   20
                  tooltiptext4    =   "Enter Fax"
                  colname5        =   "CSTEMAIL"
                  heading5        =   "eMail                                                                                   "
                  datafld5        =   "CstEMail"
                  maxlength5      =   45
                  tooltiptext5    =   "Enter eMail ID"
                  colname6        =   "CSTCPNAME"
                  heading6        =   "Contact Person                                                                  "
                  datafld6        =   "CstCpName"
                  maxlength6      =   45
                  tooltiptext6    =   "Enter Contact Person"
                  colname7        =   "CSTADD1"
                  datafld7        =   "CstAdd1"
                  recalcon7       =   "WCSTADD1"
                  maxlength7      =   45
                  style7          =   0
                  colname8        =   "CSTADD2"
                  datafld8        =   "CstAdd2"
                  recalcon8       =   "WCSTADD2"
                  maxlength8      =   45
                  style8          =   0
                  colname9        =   "CSTADD3"
                  datafld9        =   "CstAdd3"
                  recalcon9       =   "WCSTADD3"
                  maxlength9      =   45
                  style9          =   0
                  colname10       =   "CSTPRTKEY"
                  datafld10       =   "CstPrtKey"
                  maxlength10     =   1
                  style10         =   0
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   7
               Left            =   -74880
               TabIndex        =   178
               Top             =   360
               Width           =   14880
               Begin MwfCtl.MWCTL_FLX GrdCustInst 
                  Height          =   7980
                  Left            =   90
                  TabIndex        =   179
                  Top             =   90
                  Width           =   14655
                  _ExtentX        =   25850
                  _ExtentY        =   14076
                  Cols            =   4
                  colname1        =   "CIDMCTG"
                  heading1        =   "Dsg Ctg "
                  datafld1        =   "CiDmCtg"
                  datatype1       =   4
                  maxlength1      =   5
                  tooltiptext1    =   "Enter Design Category"
                  colname2        =   "CICUSTPRDINST"
                  heading2        =   $"EmrFrmCustMst.frx":00FC
                  datafld2        =   "CiCustPrdInst"
                  maxlength2      =   250
                  tooltiptext2    =   "Enter Customer Production Instruction"
                  colname3        =   "CIPRTKEY"
                  heading3        =   "Prt"
                  datafld3        =   "CiPrtKey"
                  maxlength3      =   1
                  style3          =   0
                  tooltiptext3    =   "Prt Key"
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   5
               Left            =   -74940
               TabIndex        =   84
               Top             =   660
               Width           =   15000
               Begin MwfCtl.MWCTL_FLX GrdCustRm 
                  Height          =   7980
                  Left            =   90
                  TabIndex        =   85
                  Top             =   90
                  Width           =   12435
                  _ExtentX        =   21934
                  _ExtentY        =   14076
                  Cols            =   8
                  colname1        =   "CRRMCTG"
                  heading1        =   "Ctg"
                  datafld1        =   "CrRmCtg"
                  datatype1       =   4
                  maxlength1      =   2
                  tooltiptext1    =   "Enter Rm Ctg"
                  colname2        =   "CROURRMCD"
                  heading2        =   "Our Rm Cd                               "
                  datafld2        =   "CrOurRmCd"
                  datatype2       =   4
                  maxlength2      =   16
                  tooltiptext2    =   "Enter Our Rm Code"
                  colname3        =   "CRFRLN"
                  heading3        =   "Fr Ln/Sv/Ptr "
                  datafld3        =   "CrFrLn"
                  datatype3       =   2
                  mask3           =   "####0.0000"
                  maxlength3      =   10
                  tooltiptext3    =   "Enter From Length/Sieve/Pointer"
                  colname4        =   "CRTOLN"
                  heading4        =   "To Ln/Sv/Ptr "
                  datafld4        =   "CrToLn"
                  datatype4       =   2
                  mask4           =   "####0.0000"
                  maxlength4      =   10
                  tooltiptext4    =   "Enter To Length/Sieve/Pointer"
                  colname5        =   "CRCUSTRMCD"
                  heading5        =   "Customer Rm Code             "
                  datafld5        =   "CrCustRmCd"
                  datatype5       =   4
                  maxlength5      =   20
                  tooltiptext5    =   "Enter Customer Rm Code"
                  colname6        =   "CRDESC"
                  heading6        =   "Description                                                                                   "
                  datafld6        =   "CrDesc"
                  maxlength6      =   45
                  tooltiptext6    =   "Enter Description"
                  colname7        =   "CRPRTKEY"
                  heading7        =   "Prtn "
                  datafld7        =   "CrPrtKey"
                  maxlength7      =   1
                  style7          =   0
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   4
               Left            =   -74940
               TabIndex        =   82
               Top             =   660
               Width           =   15000
               Begin MwfCtl.MWCTL_FLX GrdCustAly 
                  Height          =   7980
                  Left            =   90
                  TabIndex        =   83
                  Top             =   90
                  Width           =   6495
                  _ExtentX        =   11456
                  _ExtentY        =   14076
                  Cols            =   4
                  colname1        =   "CARMCD"
                  heading1        =   "Rm Code                                  "
                  datafld1        =   "CaRmCd"
                  datatype1       =   4
                  maxlength1      =   16
                  tooltiptext1    =   "Enter Raw Material Code"
                  colname2        =   "CADMCOL"
                  heading2        =   "Dsg Col "
                  datafld2        =   "CaDmCol"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Design Colour"
                  colname3        =   "CAALYCD"
                  heading3        =   "Alloy Code                                "
                  datafld3        =   "CaAlyCd"
                  datatype3       =   4
                  maxlength3      =   16
                  tooltiptext3    =   "Enter Alloy Code"
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   6
               Left            =   -74940
               TabIndex        =   86
               Top             =   660
               Width           =   15000
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   7440
                  Index           =   31
                  Left            =   90
                  TabIndex        =   87
                  ToolTipText     =   "Enter History Of The Customer"
                  Top             =   630
                  Width           =   14835
                  _ExtentX        =   26167
                  _ExtentY        =   13123
                  Alignment       =   3
                  MaxLength       =   0
                  DataField       =   "CmHistory"
                  IdName          =   "CMHISTORY"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "History"
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
                  Index           =   28
                  Left            =   135
                  TabIndex        =   128
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   3
               Left            =   60
               TabIndex        =   80
               Top             =   660
               Width           =   15000
               Begin MwfCtl.MWCTL_FLX GrdCustDsg 
                  Height          =   7800
                  Left            =   90
                  TabIndex        =   81
                  Top             =   90
                  Width           =   14925
                  _ExtentX        =   26326
                  _ExtentY        =   13758
                  Cols            =   7
                  colname1        =   "CDOURDMCD"
                  heading1        =   "Our Design                  "
                  datafld1        =   "CdOurDmCd"
                  datatype1       =   4
                  maxlength1      =   15
                  tooltiptext1    =   "Enter Our Design Code"
                  colname2        =   "CDOURDMSZ"
                  heading2        =   "Size      "
                  datafld2        =   "CdOurDmSz"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Our Design Size"
                  colname3        =   "CDOURSFX"
                  heading3        =   "Suffix                 "
                  datafld3        =   "CdOurSfx"
                  datatype3       =   4
                  maxlength3      =   15
                  tooltiptext3    =   "Enter Our Suffix"
                  colname4        =   "CDCUSTDMCD"
                  heading4        =   "Cust Design                                                                                                       "
                  datafld4        =   "CdCustDmCd"
                  datatype4       =   4
                  maxlength4      =   40
                  tooltiptext4    =   "Enter Customer Design Code"
                  colname5        =   "CDDESC"
                  heading5        =   $"EmrFrmCustMst.frx":01F8
                  datafld5        =   "CdDesc"
                  maxlength5      =   120
                  tooltiptext5    =   "Enter Description"
                  colname6        =   "CDPRTKEY"
                  heading6        =   "Prtn "
                  datafld6        =   "CdPrtKey"
                  maxlength6      =   1
                  style6          =   0
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8130
               Index           =   2
               Left            =   -74880
               TabIndex        =   44
               Top             =   660
               Width           =   14880
               Begin VB.Frame FraZselfDepth 
                  Height          =   1245
                  Left            =   7740
                  TabIndex        =   174
                  Top             =   1170
                  Width           =   3045
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   61
                     Left            =   2565
                     TabIndex        =   66
                     ToolTipText     =   "Enter ZSelf depth for PDCM (Y/N)"
                     Top             =   795
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmSeekZSelfPDCM"
                     IdName          =   "CMSEEKZSELFPDCM"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   63
                     Left            =   2565
                     TabIndex        =   65
                     ToolTipText     =   "Enter whether ZSelf depth to be seeked for Rate  (Y/N)"
                     Top             =   510
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmSeekZSelfRt"
                     IdName          =   "CMSEEKZSELFRT"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ZSelf depth for PDCM"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   56
                     Left            =   180
                     TabIndex        =   177
                     Top             =   795
                     Width           =   2475
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ZSelf depth for Rate "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   54
                     Left            =   180
                     TabIndex        =   176
                     Top             =   510
                     Width           =   2595
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  ZSelf Depth"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   7
                     Left            =   -30
                     TabIndex        =   175
                     Top             =   90
                     Width           =   3075
                  End
               End
               Begin VB.Frame Picture1 
                  Height          =   2535
                  Left            =   120
                  TabIndex        =   142
                  Top             =   3630
                  Width           =   11385
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   41
                     Left            =   1470
                     TabIndex        =   72
                     ToolTipText     =   "VAT Applicable (Yes/No)"
                     Top             =   795
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmVatYn"
                     IdName          =   "CMVATYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   43
                     Left            =   1470
                     TabIndex        =   71
                     ToolTipText     =   "Enter Billing Company Code"
                     Top             =   510
                     Width           =   675
                     _ExtentX        =   1191
                     _ExtentY        =   503
                     MaxLength       =   3
                     DataType        =   4
                     DataField       =   "CmBillCo"
                     IdName          =   "CMBILLCO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   44
                     Left            =   1470
                     TabIndex        =   74
                     ToolTipText     =   "CST Applicable (Yes/No)"
                     Top             =   1080
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmCstYn"
                     IdName          =   "CMCSTYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   45
                     Left            =   1470
                     TabIndex        =   76
                     ToolTipText     =   "Octori Applicable (Yes/No)"
                     Top             =   1365
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmOctroiYn"
                     IdName          =   "CMOCTROIYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   46
                     Left            =   2850
                     TabIndex        =   77
                     ToolTipText     =   "Enter Octroi Percentage"
                     Top             =   1365
                     Width           =   675
                     _ExtentX        =   1191
                     _ExtentY        =   503
                     Alignment       =   2
                     Mask            =   "#0.00"
                     MaxLength       =   5
                     DataType        =   2
                     DataField       =   "CmOctroiPer"
                     IdName          =   "CMOCTROIPER"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   151
                     Left            =   2850
                     TabIndex        =   73
                     ToolTipText     =   "Enter VAT Number"
                     Top             =   795
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "CmVATNO"
                     IdName          =   "CMVATNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   152
                     Left            =   2850
                     TabIndex        =   75
                     ToolTipText     =   "Enter CST Number"
                     Top             =   1080
                     Width           =   8340
                     _ExtentX        =   14711
                     _ExtentY        =   503
                     MaxLength       =   60
                     DataType        =   4
                     DataField       =   "CmCSTNO"
                     IdName          =   "CMCSTNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   50
                     Left            =   1470
                     TabIndex        =   78
                     ToolTipText     =   "Enter Tin No"
                     Top             =   1650
                     Width           =   3075
                     _ExtentX        =   5424
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "CmTinNo"
                     IdName          =   "CMTINNO"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   51
                     Left            =   1470
                     TabIndex        =   79
                     ToolTipText     =   "Enter Customer PAN"
                     Top             =   1935
                     Width           =   3075
                     _ExtentX        =   5424
                     _ExtentY        =   503
                     MaxLength       =   20
                     DataType        =   4
                     DataField       =   "CmPAN"
                     IdName          =   "CMPAN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "PAN"
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
                     Index           =   46
                     Left            =   180
                     TabIndex        =   155
                     Top             =   1935
                     Width           =   975
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "TIN"
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
                     Index           =   45
                     Left            =   180
                     TabIndex        =   154
                     Top             =   1650
                     Width           =   945
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST No"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   127
                     Left            =   1890
                     TabIndex        =   153
                     Top             =   1110
                     Width           =   765
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT No"
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
                     Index           =   126
                     Left            =   1890
                     TabIndex        =   152
                     Top             =   825
                     Width           =   735
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Octroi %"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   41
                     Left            =   1890
                     TabIndex        =   148
                     Top             =   1395
                     Width           =   885
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Octroi"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   40
                     Left            =   180
                     TabIndex        =   147
                     Top             =   1365
                     Width           =   1065
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Billing CoCd"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   39
                     Left            =   180
                     TabIndex        =   146
                     Top             =   510
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "VAT  "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   38
                     Left            =   180
                     TabIndex        =   145
                     Top             =   795
                     Width           =   1005
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "CST"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   36
                     Left            =   180
                     TabIndex        =   144
                     Top             =   1080
                     Width           =   1005
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   " Local Invoice Details"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   4
                     Left            =   -30
                     TabIndex        =   143
                     Top             =   90
                     Width           =   13395
                  End
               End
               Begin VB.Frame FraRtLookUp 
                  Height          =   2085
                  Left            =   3570
                  TabIndex        =   121
                  Top             =   1170
                  Width           =   3975
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   15
                     Left            =   2415
                     TabIndex        =   61
                     ToolTipText     =   "Enter Price List Customer From Whom Lab Rates Have To Be Looked Up For This Customer"
                     Top             =   795
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "CmLkUpLabRt"
                     IdName          =   "CMLKUPLABRT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   27
                     Left            =   2415
                     TabIndex        =   60
                     ToolTipText     =   "Enter Price List Customer From Whom Rm Rates Have To Be Looked Up For This Customer"
                     Top             =   510
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "CmLkUpRmRt"
                     IdName          =   "CMLKUPRMRT"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   28
                     Left            =   2415
                     TabIndex        =   62
                     ToolTipText     =   "Enter Price List Customer From Whom Metal Loss % Have To Be Looked Up For This Customer"
                     Top             =   1080
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "CmLkUpMetLs"
                     IdName          =   "CMLKUPMETLS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   70
                     Left            =   2415
                     TabIndex        =   63
                     ToolTipText     =   "Enter Price List Customer From Whom MRP Discount Have To Be Looked Up For This Customer"
                     Top             =   1365
                     Width           =   1410
                     _ExtentX        =   2487
                     _ExtentY        =   503
                     MaxLength       =   8
                     DataType        =   4
                     DataField       =   "CmLkUpMrpDisc"
                     IdName          =   "CMLKUPMRPDISC"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "MRP Discount Look Up"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   59
                     Left            =   180
                     TabIndex        =   185
                     Top             =   1365
                     Width           =   2205
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Rate Look Up"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   1
                     Left            =   -30
                     TabIndex        =   138
                     Top             =   90
                     Width           =   4005
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Metal Loss Look Up"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   25
                     Left            =   180
                     TabIndex        =   124
                     Top             =   1080
                     Width           =   2085
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Lab Rate Look Up"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   24
                     Left            =   180
                     TabIndex        =   123
                     Top             =   795
                     Width           =   1935
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rm Rate Look Up"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   13
                     Left            =   180
                     TabIndex        =   122
                     Top             =   510
                     Width           =   1965
                  End
               End
               Begin VB.Frame FraFavAvg 
                  Height          =   1305
                  Left            =   11040
                  TabIndex        =   102
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   1965
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   12
                     Left            =   1365
                     TabIndex        =   70
                     ToolTipText     =   "Enter Colour Stone Rate Favourable (F / A)"
                     Top             =   795
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmCsRtFA"
                     IdName          =   "CMCSRTFA"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   11
                     Left            =   1365
                     TabIndex        =   69
                     ToolTipText     =   "Enter Diamond Rate Favourable  (F / A)"
                     Top             =   510
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmDiaRtFA"
                     IdName          =   "CMDIARTFA"
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Favourable / Average"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Index           =   2
                     Left            =   -30
                     TabIndex        =   139
                     Top             =   90
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Dia Rt "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   10
                     Left            =   180
                     TabIndex        =   107
                     Top             =   510
                     Width           =   1155
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "ColStn Rt "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   11
                     Left            =   180
                     TabIndex        =   106
                     Top             =   795
                     Width           =   1155
                  End
               End
               Begin VB.Frame FraOrdAct 
                  Height          =   2385
                  Left            =   150
                  TabIndex        =   101
                  Top             =   1170
                  Width           =   3135
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   8
                     Left            =   2175
                     TabIndex        =   54
                     ToolTipText     =   "Enter All Weights From Order (Y / N)"
                     Top             =   540
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmIWtEqOrd"
                     ReCalcParent    =   "CMIWTFRORD"
                     IdName          =   "CMIWTEQORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   9
                     Left            =   2175
                     TabIndex        =   55
                     ToolTipText     =   "Enter Weights From Order"
                     Top             =   825
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   503
                     MaxLength       =   5
                     DataType        =   4
                     DataField       =   "CmIWtFrOrd"
                     ReCalcOn        =   "CMIWTEQORD"
                     IdName          =   "CMIWTFRORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   10
                     Left            =   2175
                     TabIndex        =   56
                     ToolTipText     =   "Enter Invoice Gold Rate From Order (Y/N)"
                     Top             =   1110
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmIGldRtEqOrd"
                     IdName          =   "CMIGLDRTEQORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   49
                     Left            =   2175
                     TabIndex        =   57
                     ToolTipText     =   "Enter Invoice Labor Weight From Order (Y/N)"
                     Top             =   1395
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmILabWtFrOrd"
                     IdName          =   "CMILABWTFRORD"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   65
                     Left            =   2175
                     TabIndex        =   58
                     ToolTipText     =   "Enter Invoice - Wait For All Bags  (Y/N)"
                     Top             =   1680
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmInvAllBags"
                     IdName          =   "CMINVALLBAGS"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   71
                     Left            =   2175
                     TabIndex        =   59
                     ToolTipText     =   "Enter Invoice By Bag (Y/N)"
                     Top             =   1965
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CmInvByBagYN"
                     IdName          =   "CMINVBYBAGYN"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice By Bag"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   61
                     Left            =   180
                     TabIndex        =   186
                     Top             =   1965
                     Width           =   1575
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv-Wait For All Bags"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   57
                     Left            =   180
                     TabIndex        =   181
                     Top             =   1680
                     Width           =   1995
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Labor Wt From Ord"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   44
                     Left            =   180
                     TabIndex        =   151
                     Top             =   1395
                     Width           =   1875
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Wt From Order / Actuals"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   3
                     Left            =   -30
                     TabIndex        =   140
                     Top             =   90
                     Width           =   3105
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "All Wts From Ord "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   7
                     Left            =   180
                     TabIndex        =   105
                     Top             =   540
                     Width           =   1755
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Wt From Ord "
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   8
                     Left            =   180
                     TabIndex        =   104
                     Top             =   825
                     Width           =   1755
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Gld Rt From Ord"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   9
                     Left            =   180
                     TabIndex        =   103
                     Top             =   1110
                     Width           =   1725
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   14
                  Left            =   4425
                  TabIndex        =   48
                  ToolTipText     =   "Enter Gold As"
                  Top             =   270
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmGldAs"
                  IdName          =   "CMGLDAS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   7
                  Left            =   6945
                  TabIndex        =   50
                  ToolTipText     =   "Enter Fixed Price (Y / N)"
                  Top             =   315
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "CMFixPrc"
                  IdName          =   "CMFIXPRC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   67
                  Left            =   9600
                  TabIndex        =   52
                  ToolTipText     =   "Enter Option By Which Rounding Off Takes Place For Order Sales Price"
                  Top             =   555
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmHRndOrdSalPrc"
                  IdName          =   "CMHRNDORDSALPRC"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   6
                  Left            =   2565
                  TabIndex        =   45
                  ToolTipText     =   "Enter Multiplying Factor"
                  Top             =   270
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "CmMulBy"
                  ReCalcParent    =   "CMORDMULBYYN"
                  IdName          =   "CMMULBY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   33
                  Left            =   9600
                  TabIndex        =   51
                  ToolTipText     =   "Enter Currency Code"
                  Top             =   270
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmCurCd"
                  IdName          =   "CMCURCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   42
                  Left            =   4425
                  TabIndex        =   49
                  ToolTipText     =   "Enter Labour As"
                  Top             =   555
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmLabAs"
                  IdName          =   "CMLABAS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   60
                  Left            =   2565
                  TabIndex        =   47
                  ToolTipText     =   "Enter MRP Multiplying Factor"
                  Top             =   855
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "CmMrpMulBy"
                  IdName          =   "CMMRPMULBY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   72
                  Left            =   2565
                  TabIndex        =   46
                  ToolTipText     =   "Enter Apply Multiply Factor to Invoice (Y/N)"
                  Top             =   555
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "CmOrdMulByYN"
                  ReCalcOn        =   "CMMULBY"
                  IdName          =   "CMORDMULBYYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   74
                  Left            =   9600
                  TabIndex        =   53
                  ToolTipText     =   "Enter Apply TCS (Y / N)"
                  Top             =   840
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "CmTCSYN"
                  IdName          =   "CMTCSYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   73
                  Left            =   5985
                  TabIndex        =   64
                  ToolTipText     =   "Enter Customer From Whom PDCM has to be Looked Up For This Customer"
                  Top             =   3270
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmLkUpPDCM"
                  IdName          =   "CMLKUPPDCM"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   75
                  Left            =   10290
                  TabIndex        =   67
                  ToolTipText     =   "Enter N to stop taking New Orders"
                  Top             =   2430
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "CMAllwNewSOYN"
                  ReCalcParent    =   "CMALLWNEWSODESC"
                  IdName          =   "CMALLWNEWSOYN"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   855
                  Index           =   76
                  Left            =   10290
                  TabIndex        =   68
                  ToolTipText     =   "Enter reason to stop taking new Orders"
                  Top             =   2730
                  Width           =   4545
                  _ExtentX        =   8017
                  _ExtentY        =   1508
                  MaxLength       =   60
                  DataField       =   "CMALLWNEWSODESC"
                  ReCalcOn        =   "CMALLWNEWSOYN"
                  IdName          =   "CMALLWNEWSODESC"
                  Data            =   "aabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijbcdefghij"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Allow New SO Description"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Index           =   67
                  Left            =   7860
                  TabIndex        =   204
                  Top             =   2760
                  Width           =   2325
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Allow New SO "
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
                  Index           =   66
                  Left            =   7860
                  TabIndex        =   203
                  Top             =   2430
                  Width           =   2325
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PDCM Look Up"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   63
                  Left            =   3720
                  TabIndex        =   189
                  Top             =   3270
                  Width           =   1965
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Apply TCS"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   64
                  Left            =   7830
                  TabIndex        =   188
                  Top             =   885
                  Width           =   1500
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Apply Mul Fctr to Invoice"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   62
                  Left            =   210
                  TabIndex        =   187
                  Top             =   555
                  Width           =   2415
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MRP Multiply By"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   53
                  Left            =   210
                  TabIndex        =   173
                  Top             =   855
                  Width           =   1755
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lab As"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   37
                  Left            =   3510
                  TabIndex        =   141
                  Top             =   600
                  Width           =   1125
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Currency Cd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   30
                  Left            =   7800
                  TabIndex        =   129
                  Top             =   315
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Multiply By"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   5
                  Left            =   210
                  TabIndex        =   127
                  Top             =   270
                  Width           =   1275
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Round Sal Prc"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   65
                  Left            =   7830
                  TabIndex        =   126
                  Top             =   600
                  Width           =   1500
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gold As"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   14
                  Left            =   3510
                  TabIndex        =   109
                  Top             =   270
                  Width           =   1125
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fixed Prc"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   6
                  Left            =   5790
                  TabIndex        =   108
                  Top             =   315
                  Width           =   1065
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8370
               Index           =   1
               Left            =   -74940
               TabIndex        =   32
               Top             =   420
               Width           =   15000
               Begin VB.Frame Frame1 
                  Height          =   2475
                  Left            =   0
                  TabIndex        =   168
                  Top             =   2520
                  Width           =   12855
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   59
                     Left            =   1725
                     TabIndex        =   41
                     ToolTipText     =   "Enter Whether by default Show Net Wt Should be 'Y' in Invoice Printing Report"
                     Top             =   480
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CMINVSHOWGRSWTYN"
                     IdName          =   "CMINVSHOWGRSWTYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   285
                     Index           =   62
                     Left            =   4200
                     TabIndex        =   42
                     ToolTipText     =   "Enter Whether by default Show Dia Wt Should be 'Y' in Invoice Printing Report"
                     Top             =   480
                     Width           =   315
                     _ExtentX        =   556
                     _ExtentY        =   503
                     DataType        =   4
                     DataField       =   "CMINVSHOWDIAWTYN"
                     IdName          =   "CMINVSHOWDIAWTYN"
                  End
                  Begin MwfCtl.MWCTL_MED atxt 
                     Height          =   1545
                     Index           =   90
                     Left            =   1725
                     TabIndex        =   43
                     ToolTipText     =   "Enter Customer Invoice Footer Remarks"
                     Top             =   780
                     Width           =   11130
                     _ExtentX        =   19632
                     _ExtentY        =   2725
                     Alignment       =   3
                     MaxLength       =   360
                     DataField       =   "CMINVFTREM"
                     IdName          =   "CMINVFTREM"
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Inv Footer Rem"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   76
                     Left            =   80
                     TabIndex        =   172
                     Top             =   780
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Dia/CS Wt"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   55
                     Left            =   2520
                     TabIndex        =   171
                     Top             =   480
                     Width           =   1635
                  End
                  Begin VB.Label ALBL 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Show Net Wt"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   52
                     Left            =   80
                     TabIndex        =   170
                     Top             =   480
                     Width           =   1515
                  End
                  Begin VB.Label LblMain 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "  Defaults for Invoicing Reports"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   6
                     Left            =   -30
                     TabIndex        =   169
                     Top             =   90
                     Width           =   12885
                  End
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   690
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   33
                  ToolTipText     =   "Enter Production Instruction"
                  Top             =   150
                  Width           =   11265
                  _ExtentX        =   19870
                  _ExtentY        =   1217
                  Alignment       =   3
                  MaxLength       =   250
                  DataField       =   "CmPrdInst"
                  IdName          =   "CMPRDINST"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   4
                  Left            =   3300
                  TabIndex        =   36
                  ToolTipText     =   "Enter Stamping Instruction"
                  Top             =   1200
                  Width           =   9660
                  _ExtentX        =   17039
                  _ExtentY        =   503
                  MaxLength       =   90
                  DataField       =   "CmStmpInst"
                  ReCalcOn        =   "WCMSTMPCD"
                  IdName          =   "CMSTMPINST"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   5
                  Left            =   1695
                  TabIndex        =   37
                  ToolTipText     =   "Enter Delivery Terms"
                  Top             =   1485
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmDelTerms"
                  IdName          =   "DMDELTERMS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   16
                  Left            =   3300
                  TabIndex        =   39
                  ToolTipText     =   "Enter Pay Terms"
                  Top             =   1770
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmPayTerms"
                  ReCalcOn        =   "CMPAYTERMSCD"
                  IdName          =   "DMPAYTERMS"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   40
                  Left            =   1695
                  TabIndex        =   40
                  ToolTipText     =   "Enter Vessel"
                  Top             =   2115
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "CmVessel"
                  IdName          =   "CMVESSEL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   48
                  Left            =   1695
                  TabIndex        =   35
                  ToolTipText     =   "Enter Stamping Code"
                  Top             =   1200
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  ReCalcParent    =   "CMSTMPINST"
                  IdName          =   "WCMSTMPCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   64
                  Left            =   1695
                  TabIndex        =   34
                  ToolTipText     =   "Enter Special remarks"
                  Top             =   840
                  Width           =   11265
                  _ExtentX        =   19870
                  _ExtentY        =   503
                  MaxLength       =   100
                  DataField       =   "CmOmSplRem"
                  IdName          =   "CMOMSPLREM"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   69
                  Left            =   1695
                  TabIndex        =   38
                  ToolTipText     =   "Enter Pay Term Code"
                  Top             =   1800
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataField       =   "CmPayTermsCd"
                  ReCalcParent    =   "DMPAYTERMS"
                  IdName          =   "CMPAYTERMSCD"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   " -- "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   60
                  Left            =   3000
                  TabIndex        =   184
                  Top             =   1800
                  Width           =   405
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Spl Remarks"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   33
                  Left            =   60
                  TabIndex        =   180
                  Top             =   840
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   " -- "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   43
                  Left            =   3000
                  TabIndex        =   150
                  Top             =   1200
                  Width           =   405
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vessel"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   35
                  Left            =   60
                  TabIndex        =   136
                  Top             =   2115
                  Width           =   1305
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pay Terms"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   15
                  Left            =   60
                  TabIndex        =   110
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Del Terms"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   4
                  Left            =   60
                  TabIndex        =   100
                  Top             =   1485
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Stamping Instr"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   2
                  Left            =   60
                  TabIndex        =   99
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Instr"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   1
                  Left            =   60
                  TabIndex        =   98
                  Top             =   150
                  Width           =   1545
               End
            End
            Begin VB.Frame FraTabCustDet 
               BorderStyle     =   0  'None
               Height          =   8010
               Index           =   0
               Left            =   -74880
               TabIndex        =   5
               Top             =   360
               Width           =   14880
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   17
                  Left            =   1935
                  TabIndex        =   6
                  ToolTipText     =   "Enter Address"
                  Top             =   195
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmAdd1"
                  IdName          =   "CMADD1"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   18
                  Left            =   1935
                  TabIndex        =   7
                  ToolTipText     =   "Enter Address"
                  Top             =   480
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmAdd2"
                  IdName          =   "CMADD2"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   19
                  Left            =   1935
                  TabIndex        =   8
                  ToolTipText     =   "Enter Address"
                  Top             =   765
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmAdd3"
                  IdName          =   "CMADD3"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   20
                  Left            =   1935
                  TabIndex        =   9
                  ToolTipText     =   "Enter Telephone No"
                  Top             =   1050
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "CmTel"
                  IdName          =   "CMTEL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   21
                  Left            =   1935
                  TabIndex        =   10
                  ToolTipText     =   "Enter Fax No"
                  Top             =   1335
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "CmFax"
                  IdName          =   "CMFAX"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   22
                  Left            =   1935
                  TabIndex        =   11
                  ToolTipText     =   "Enter E-Mail"
                  Top             =   1620
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmeMail"
                  IdName          =   "CMEMAIL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   23
                  Left            =   2895
                  TabIndex        =   18
                  ToolTipText     =   "Enter Port Of Discharge"
                  Top             =   3330
                  Width           =   4105
                  _ExtentX        =   7250
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "CmPortofDisch"
                  ReCalcOn        =   "CMPOD"
                  IdName          =   "CMPORTOFDISCH"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   24
                  Left            =   2895
                  TabIndex        =   20
                  ToolTipText     =   "Enter Final Destination"
                  Top             =   3615
                  Width           =   4105
                  _ExtentX        =   7250
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   30
                  DataType        =   4
                  DataField       =   "CmFinDest"
                  ReCalcOn        =   "CMFINDSTCD"
                  IdName          =   "CMFINDEST"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   25
                  Left            =   1935
                  TabIndex        =   21
                  ToolTipText     =   "Enter Buyer"
                  Top             =   3900
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataType        =   4
                  DataField       =   "CmBuyer"
                  IdName          =   "CMBUYER"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   26
                  Left            =   1935
                  TabIndex        =   12
                  ToolTipText     =   "Enter Contact Person Name"
                  Top             =   1905
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmCpName"
                  IdName          =   "CMCPNAME"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   30
                  Left            =   1935
                  TabIndex        =   25
                  ToolTipText     =   "Enter Insurance By"
                  Top             =   5040
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataField       =   "CmHInsBy"
                  IdName          =   "CMHINSBY"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   32
                  Left            =   1935
                  TabIndex        =   14
                  ToolTipText     =   "Enter Code for the Associated Sales Executive"
                  Top             =   2475
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmSalPer"
                  IdName          =   "CMSALPER"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   29
                  Left            =   1935
                  TabIndex        =   13
                  ToolTipText     =   "Enter Customer Region Code"
                  Top             =   2190
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmRegnCd"
                  IdName          =   "CMREGNCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   34
                  Left            =   1935
                  TabIndex        =   16
                  ToolTipText     =   "Enter Agent Code"
                  Top             =   3045
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmAgCd"
                  IdName          =   "CMAGCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   35
                  Left            =   1935
                  TabIndex        =   15
                  ToolTipText     =   "Enter Bank  Code"
                  Top             =   2760
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmBkCd"
                  IdName          =   "CMBKCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   36
                  Left            =   1935
                  TabIndex        =   22
                  ToolTipText     =   "Enter Buyer's Address"
                  Top             =   4185
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmBuyAdd1"
                  IdName          =   "CMBUYADD1"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   37
                  Left            =   1935
                  TabIndex        =   23
                  ToolTipText     =   "Enter  Buyer's  Address"
                  Top             =   4470
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmBuyAdd2"
                  IdName          =   "CMBUYADD2"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   38
                  Left            =   1935
                  TabIndex        =   24
                  ToolTipText     =   "Enter  Buyer's  Address"
                  Top             =   4755
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "CmBuyAdd3"
                  IdName          =   "CMBUYADD3"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   47
                  Left            =   1935
                  TabIndex        =   26
                  ToolTipText     =   "Enter Bank  Code"
                  Top             =   5325
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "CmOurBkCd"
                  IdName          =   "CMOURBKCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   55
                  Left            =   1935
                  TabIndex        =   27
                  ToolTipText     =   "Enter Bank  Code"
                  Top             =   5640
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmGrp"
                  IdName          =   "CMGRP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   56
                  Left            =   1935
                  TabIndex        =   28
                  ToolTipText     =   "Enter Bank  Code"
                  Top             =   5950
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmTyp"
                  IdName          =   "CMTYP"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   57
                  Left            =   1935
                  TabIndex        =   17
                  ToolTipText     =   "Enter Code for the Associated Sales Executive"
                  Top             =   3330
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   503
                  MaxLength       =   6
                  DataType        =   4
                  DataField       =   "CmPod"
                  ReCalcParent    =   "CMPORTOFDISCH"
                  IdName          =   "CMPOD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   58
                  Left            =   1935
                  TabIndex        =   19
                  ToolTipText     =   "Enter Code for the Associated Sales Executive"
                  Top             =   3615
                  Width           =   500
                  _ExtentX        =   873
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataType        =   4
                  DataField       =   "CmFinDstCd"
                  ReCalcParent    =   "CMFINDEST"
                  IdName          =   "CMFINDSTCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   66
                  Left            =   1935
                  TabIndex        =   30
                  ToolTipText     =   "Enter The Valid Co Cd's For Editing this Record"
                  Top             =   6525
                  Width           =   8040
                  _ExtentX        =   14182
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataType        =   4
                  DataField       =   "CmValidCoCd"
                  IdName          =   "CMVALIDCOCD"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   68
                  Left            =   1935
                  TabIndex        =   29
                  ToolTipText     =   "Enter Cell Name"
                  Top             =   6240
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "CmCell"
                  IdName          =   "CMCELL"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   77
                  Left            =   1935
                  TabIndex        =   31
                  ToolTipText     =   "Enter Customer Created Date"
                  Top             =   6840
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "CmCreateDt"
                  IdName          =   "CMCREATEDT"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Created Date"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   71
                  Left            =   90
                  TabIndex        =   215
                  Top             =   6840
                  Width           =   1665
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cell"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   58
                  Left            =   90
                  TabIndex        =   183
                  Top             =   6270
                  Width           =   1665
               End
               Begin VB.Label LblCmValidCoCd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Valid Co Cds"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   90
                  TabIndex        =   182
                  Top             =   6555
                  Width           =   1245
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Type"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   51
                  Left            =   90
                  TabIndex        =   167
                  Top             =   5950
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Group"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   50
                  Left            =   90
                  TabIndex        =   166
                  Top             =   5640
                  Width           =   1575
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Our Bank Code"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   42
                  Left            =   90
                  TabIndex        =   149
                  Top             =   5325
                  Width           =   1575
               End
               Begin VB.Label albl_BuyAdd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Buyer's Address"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   90
                  TabIndex        =   134
                  Top             =   4185
                  Width           =   1725
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agent Code"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   32
                  Left            =   90
                  TabIndex        =   133
                  Top             =   3045
                  Width           =   1815
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Code"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   31
                  Left            =   90
                  TabIndex        =   132
                  Top             =   2760
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Executive"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   29
                  Left            =   90
                  TabIndex        =   131
                  Top             =   2475
                  Width           =   1965
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Region"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   26
                  Left            =   90
                  TabIndex        =   130
                  Top             =   2190
                  Width           =   825
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Insurance By"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   27
                  Left            =   90
                  TabIndex        =   125
                  Top             =   5010
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cont Pers Name"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   23
                  Left            =   90
                  TabIndex        =   118
                  Top             =   1905
                  Width           =   1695
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Buyer (Invoice)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   22
                  Left            =   90
                  TabIndex        =   117
                  Top             =   3900
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Final Destination"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   21
                  Left            =   90
                  TabIndex        =   116
                  Top             =   3615
                  Width           =   1785
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Port Of Destination"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   20
                  Left            =   90
                  TabIndex        =   115
                  Top             =   3330
                  Width           =   1845
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-Mail"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   19
                  Left            =   90
                  TabIndex        =   114
                  Top             =   1620
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "GST"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   18
                  Left            =   90
                  TabIndex        =   113
                  Top             =   1335
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   17
                  Left            =   90
                  TabIndex        =   112
                  Top             =   1050
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   16
                  Left            =   90
                  TabIndex        =   111
                  Top             =   195
                  Width           =   1545
               End
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   2
            ToolTipText     =   "Enter Customer Name"
            Top             =   300
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   503
            MaxLength       =   45
            DataField       =   "CMName"
            IdName          =   "CMNAME"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   39
            Left            =   10080
            TabIndex        =   3
            ToolTipText     =   "Enter Valid (Y/N)"
            Top             =   270
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "CMVALIDYN"
            IdName          =   "CMVALIDYN"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Valid (Y/N)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   34
            Left            =   8790
            TabIndex        =   135
            Top             =   270
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   97
            Top             =   290
            Width           =   855
         End
      End
      Begin VB.Frame FraCpy 
         Height          =   1155
         Left            =   5820
         TabIndex        =   89
         Top             =   8010
         Width           =   3750
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   495
            Left            =   3000
            TabIndex        =   91
            Top             =   540
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
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
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   13
            Left            =   1440
            TabIndex        =   90
            ToolTipText     =   "Enter Source Customer Code"
            Top             =   525
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCUSTFR"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy Alloy Details From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   -30
            TabIndex        =   137
            Top             =   90
            Width           =   3780
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   119
            Top             =   525
            Width           =   1395
         End
      End
      Begin VB.Frame FraImpExc 
         Height          =   1515
         Left            =   3180
         TabIndex        =   191
         Top             =   7680
         Width           =   8235
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   248
            Left            =   660
            TabIndex        =   192
            ToolTipText     =   "Excel File Name and Path"
            Top             =   555
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   250
            IdName          =   "WXLFILE"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   249
            Left            =   750
            TabIndex        =   193
            ToolTipText     =   "Specify the Excel Sheet Number"
            Top             =   1065
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WXLSHTNO"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   250
            Left            =   2220
            TabIndex        =   194
            ToolTipText     =   "Specify Excel From Row Number"
            Top             =   1065
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WXLFRROW"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   251
            Left            =   3870
            TabIndex        =   195
            ToolTipText     =   "Specify Excel To Row Number"
            Top             =   1065
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WXLTOROW"
         End
         Begin MwfCtl.MWCTL_BTN1 CmdImpExcGo 
            Height          =   495
            Left            =   7560
            TabIndex        =   196
            ToolTipText     =   "Click Here To Copy"
            Top             =   960
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
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
         Begin MwfCtl.MWCTL_BTN1 CmdImpExcPath 
            Height          =   495
            Left            =   7560
            TabIndex        =   197
            ToolTipText     =   "Click Here To Copy"
            Top             =   450
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   873
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Imp"
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Import From Excel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   11
            Left            =   -30
            TabIndex        =   202
            Top             =   120
            Width           =   5640
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   201
            Left            =   30
            TabIndex        =   201
            Top             =   555
            Width           =   615
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Sht No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   202
            Left            =   30
            TabIndex        =   200
            Top             =   1065
            Width           =   765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Fr Rw"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   203
            Left            =   1560
            TabIndex        =   199
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "To Rw"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   204
            Left            =   3030
            TabIndex        =   198
            Top             =   1065
            Width           =   645
         End
      End
   End
   Begin VB.Frame FraCpyRm 
      Height          =   1395
      Left            =   4440
      TabIndex        =   157
      Top             =   7800
      Width           =   5055
      Begin MwfCtl.MWCTL_BTN1 CmdCpyRMGo 
         Height          =   495
         Left            =   4370
         TabIndex        =   162
         Top             =   660
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
         BackColor       =   -2147483633
         MaskColor       =   8454143
         Caption         =   "&Go"
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
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   52
         Left            =   1440
         TabIndex        =   158
         ToolTipText     =   "Enter Source Customer Code"
         Top             =   525
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WCUSTRMFR"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   53
         Left            =   1440
         TabIndex        =   159
         ToolTipText     =   "Enter from RM Ctg"
         Top             =   840
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WFRRMCTG"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   54
         Left            =   3720
         TabIndex        =   160
         ToolTipText     =   "Enter to RM Ctg"
         Top             =   840
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WTORMCTG"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "To RmCtg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   49
         Left            =   2640
         TabIndex        =   165
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr RmCtg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   48
         Left            =   120
         TabIndex        =   164
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Source Cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   47
         Left            =   120
         TabIndex        =   163
         Top             =   525
         Width           =   1395
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Copy RM Details From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   -30
         TabIndex        =   161
         Top             =   90
         Width           =   5445
      End
   End
   Begin MSComDlg.CommonDialog CdImpExc 
      Left            =   2910
      Top             =   9300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmrFrmCustMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----   General Comments on the Entry Form  -------'
'
'   Three Menus 'C'ustomer Master, 'S'upplier, 'P'rice List
'   Adc(CmCtg), invisible, in the Key Frame
'   Tab Page, TabCustDet with 5 Tabs
'     1.Customer Detail   2. Instructions
'     3.Factors, 2 frames within for display purpose only.
'     4.Cust Design, Grid GrdCustDsg
'     5.Cust Rm , Grid GrdCustRm
'     6.Cust Alloy, Grid GrdCustAly with CmdCpy button to Copy into grid.
'   FraCpy on the form for Copy of GrdCustAly, CmdCpyGo Button.
'   X - CustMst,  A- CustDsg, D-CustAly.
'   private sub CpyAlyDet for Copying Alloy details into grid.
'
'---- vbvb

Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_MultiCoMod As Boolean
'*** For the Toggle Options of Command Buttons
Enum en_CustFra
  Cpy = 0
  CpyRm = 1
  ImpExc = 2 'Uni.9- Import from Excel
  Scope = 3   '6.1-90
  'Sachin 4.1.0.0
End Enum
'*** Jay 2.14 *** (01/03/07)
Dim ms_UsrLvl As Integer
'*** Jay 2.14 *** (01/03/07)
Dim ms_ChkQtsIdName As String

Private Sub ADC_Load()
  '*** Set the First Key Control to Customer Code
  '*** Set the First Non Key Control to Customer Name
  '*** Set the Child Property Of ADC to GrdCustDsg,GrdCustAly
      '(i.e. both GrdCustDsg and GrdCustAly are children of ADC)
  '*** Set the Previous and Next Control Properties for ADC and the Grids GrdCustDsg, GrdCustRm and GrdCustAly
  '*** Set the frame FraNKeyAll as the top most frame
  '*** Set the current tab as 0 (This is done so that whatever be the active tab during
      'design time, the first tab is set as the active tab at run time)
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
    
  If ws_MultiCoMod And (UCase(ADC.MenuCd) = UCase("CustMst") Or UCase(ADC.MenuCd) = UCase("SuppMst") Or UCase(ADC.MenuCd) = UCase("PLMst")) Then _
     ADC("CmValidCoCd").Enabled = True: ADC("CmValidCoCd").Visible = True: LblCmValidCoCd.Visible = True
  If Not ws_MultiCoMod Or (UCase(ADC.MenuCd) <> UCase("CustMst") And UCase(ADC.MenuCd) <> UCase("SuppMst") And UCase(ADC.MenuCd) <> UCase("PLMst")) Then _
     ADC("CmValidCoCd").Enabled = False: ADC("CmValidCoCd").Visible = False: LblCmValidCoCd.Visible = False
  
  Set ADC.FirKeyCtl = ADC("CmCd")
  Set ADC.FirNKeyCtl = ADC("CmName")
           
  ADC.Child = "GrdCustDsg,GrdCustAly,GrdCustRm,GrdCustInst,GrdCustStore"  '578-20 GrdCustStore added
  Set GrdCustDsg.PrevCtl = ADC
  Set GrdCustDsg.NextCtl = ADC
  Set GrdCustRm.PrevCtl = ADC
  Set GrdCustRm.NextCtl = ADC
  Set GrdCustAly.PrevCtl = ADC
  Set GrdCustAly.NextCtl = ADC
  Set GrdCustInst.PrevCtl = ADC
  Set GrdCustInst.NextCtl = ADC
  Set GrdCustStore.PrevCtl = ADC  '578-20
  Set GrdCustStore.NextCtl = ADC  '578-20

  FraNKeyAll.ZOrder (0)
  TabCustDet.Tab = 0
  '*** jay 2.14 *** [01/03/07]
  ms_UsrLvl = moCn.GetFldVal("Select PNum from Param where Ptyp = 'USR' and PmCd = '" + gs_UsrCd + "'")
  '*** jay 2.14 *** [01/03/07]
  '*** Jen
'  App.HelpFile = "F:\Jen\JenVBHelp\Emr6Hlp1.Hlp"
'  Me.HelpContextID = 21
  '*** Jen
  
  '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  GrdCustDsg.ColProp("CdPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdCustRm.ColProp("CrPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdCustInst.ColProp("CiPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
    If ADC("CmCtg") = "T" Or ADC("CmCtg") = "S" Or ADC("CmCtg") = "P" Then
      ADC("CMALLWNEWSOYN").Visible = False
      ADC("CMALLWNEWSOYNDESC").Visible = False
    End If
End Sub
Private Sub ADC_KeyWhen()
  '*** Set the caption for the Form based on the Menu Code selected
  '*** If Menu Code is CustMst the Caption= 'Customer Master'
  '*** If Menu Code is SuppMst the Caption= 'Supplier Master'
  '*** If Menu Code is PLMst the Caption= 'Price List Customer Master'
  '*** Disable all the command buttons (Because since the focus is on the key frame and
      'the user accidently clicks any of the command buttons then it might give a
      'run time error. So it is better to disable all the command buttons when the focus is
      'in the Key Frame.
  '*** Hide all frames except the frame FraNKeyAll

  If UCase(ADC.MenuCd) = UCase("CustMst") Then Me.Caption = "Customer Master": ADC("CmCtg") = "C"
  If UCase(ADC.MenuCd) = UCase("SuppMst") Then Me.Caption = "Supplier Master": ADC("CmCtg") = "S"
  If UCase(ADC.MenuCd) = UCase("PLMst") Then Me.Caption = "Price List Customer Master": ADC("CmCtg") = "P"
  If UCase(ADC.MenuCd) = UCase("CoMst") Then Me.Caption = "Company Master": ADC("CmCtg") = "Z"
  If UCase(ADC.MenuCd) = UCase("ContractorMst") Then Me.Caption = "Contractor Master": ADC("CmCtg") = "T"
  If UCase(ADC.MenuCd) = UCase("CtlgMst") Then Me.Caption = "Catalogue Master": ADC("CmCtg") = "CC"   'Geeta***Crm**2.12
  
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/04/2004
  If UCase$(ADC("CmCtg")) = "P" Or UCase$(ADC("CmCtg")) = "Z" Then
    ALBL(34).Visible = False: ADC("CmValidYn").Visible = False: ADC("CmValidYN") = "Y"
  End If
' ###########################################  Manoj  ###########################################

  ' Emr 212
  If UCase(ADC.MenuCd) = UCase("CtlgMst") Then
    TabCustDet.TabVisible(0) = False: TabCustDet.TabVisible(1) = False
    TabCustDet.TabVisible(2) = False: TabCustDet.TabVisible(4) = False
    TabCustDet.TabVisible(5) = False: TabCustDet.TabVisible(6) = False
    TabCustDet.TabCaption(3) = "<&1> Cust Design"
    CmdCpy.Visible = False
    CmdCpyRm.Visible = False    'Sachin - 4.1.0.0
    CmdImpExc.Visible = False
  End If
  ' Emr 212

  Call EnaDisaCmds(True)
  Call HideAllFras
  
  '*** jay 2.14 ***
  Dim wCnt As Integer
  '*** jay 2.14 *** [01/03/07]
  'If moCn.GetFldVal("Select PNum from Param where Ptyp = 'USR' and PmCd = '" + gs_UsrCd + "'") > 5 Then
  If ms_UsrLvl > 5 Then
  '*** jay 2.14 *** [01/03/07]
    ADC.AllowAdd = True
    For wCnt = 0 To TabCustDet.Tabs - 1
      TabCustDet.TabVisible(wCnt) = False
      TabCustDet.TabEnabled(wCnt) = False
      FraTabCustDet(wCnt).Visible = False
      FraTabCustDet(wCnt).Enabled = False
    Next
    TabCustDet.TabVisible(3) = True
    TabCustDet.TabEnabled(3) = True
    TabCustDet.TabCaption(3) = "<&1> Cust Design"
    FraTabCustDet(3).Visible = True
    FraTabCustDet(3).Enabled = True
    TabCustDet.Tab = 3
  End If
  '*** jay 2.14 ***
  'stellar.1- IdName of text boxes to be check quotes
  ms_ChkQtsIdName = ADC("CmName").IdName + "|" + ADC("CmAdd1").IdName + "|" + ADC("CmAdd2").IdName + "|" + ADC("CmAdd3").IdName + "|" + ADC("CmTel").IdName + "|" + _
    ADC("CmFax").IdName + "|" + ADC("CmeMail").IdName + "|" + ADC("CmCpName").IdName + "|" + ADC("CmBuyer").IdName + "|" + _
    ADC("CmBuyAdd1").IdName + "|" + ADC("CmBuyAdd2").IdName + "|" + ADC("CmBuyAdd3").IdName + "|" + ADC("CmHInsBy").IdName + "|" + ADC("CmPrdInst").IdName + "|" + _
    ADC("CmOmSplRem").IdName + "|" + ADC("CmStmpInst").IdName + "|" + ADC("DmDelTerms").IdName + "|" + ADC("DmPayTerms").IdName + "|" + ADC("CmVessel").IdName + "|" + _
    ADC("CmInvFtRem").IdName + "|" + ADC("CmVATNo").IdName + "|" + ADC("CmCSTNo").IdName + "|" + ADC("CmTinNo").IdName + "|" + ADC("CmPAN").IdName + "|" + ADC("CmHistory").IdName
    CmdSC.Enabled = True    '6.1-90
    
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
'Indigo.1- Checking user access permission for a given customer
If ws_MultiCoMod Then
  Dim wErrMsg As String
  If Not ChkCmpUsrAccess(ADC("CmCtg"), ADC.MenuCd, "", "", ADC("CmCd"), wErrMsg) Then
    pr_Cancel = True: pr_ErrMsg = "Customer/Supplier/PL already exists, you cannot add again and you don't have permissions to view the details": Exit Sub
  End If
End If

  If moCn.GetFldVal("Select HCtbMod From Head where HCoCd= '" + gs_CoCd + "' and " + _
     "HCd= '" + ctSelfCmCd + "'") = "N" And ADC("CmCtg") = "T" Then _
     pr_Cancel = True: pr_ErrMsg = "Cannot Enter Contractor Code As Contractor Module Option Is Not Selected In Head File"
End Sub
'578-20 once address 1,2,3 in footer changed then updating to grid
Private Sub CmdAddrOk_Click()
  With GrdCustStore
    If .Rows > 1 Then
      If .Value(.Row, "CstAdd1") <> TXTADDR(0).TxtText Then .Value(.Row, "CstAdd1") = TXTADDR(0).TxtText
      If .Value(.Row, "CstAdd2") <> TXTADDR(1).TxtText Then .Value(.Row, "CstAdd2") = TXTADDR(1).TxtText
      If .Value(.Row, "CstAdd3") <> TXTADDR(2).TxtText Then .Value(.Row, "CstAdd3") = TXTADDR(2).TxtText
    End If
  End With
  ADC.AllowSave = True
  ADC.AllowDelete = True
  FraNKeyAll.Enabled = True
  GrdCustStore.Col = 1
  GrdCustStore.SetFocus

End Sub

'Uni.9- Checking excel file name, Sheet no., From & To row nos. If all correct then call Import excel funtion
Private Sub CmdImpExcGo_Click()
  If Trim(ADC("WXLFILE")) = "" Then DispMsg "Please Specify a Proper Excel File Name !!", etError: Exit Sub
  If ADC("WXLSHTNO") = 0 Or ADC("WXLFRROW") = 0 Or ADC("WXLTOROW") = 0 Then DispMsg "Please Specify a Proper Sheet No./From Row/To Row !!", etError: Exit Sub
  '578.19 when Rm tab enabled it will call ImportExcelRm
  If TabCustDet.Tab = 5 Then
    Call ImportExcelRm
  Else
    Call ImportExcel
  End If
  Call DispFra(ImpExc)
  CmdSC.Enabled = False   '6.1-90 disable scope button
End Sub
'Uni.9-Store data from excel file to Grid
Private Sub ImportExcel()
    On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wCdOurDmCd As String, wCdOurDmSz As String, wCdOurSfx As String, wCdCustDmCd As String, wCdDesc As String
            Dim wCdCtg As String, wCdCd As String

    'Checks proper excel file format
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Checking proper sheet no.
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    
    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select CdCtg ,CdCd ,CdOurDmCd ,CdCustDmCd, CdDesc, CdOurDmSz, CdOurSfx Into #TmpXls From CustDsg Where 1=2"
            
    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        'Checking Customer code
        wCdCtg = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 1)))
        
        wCdCd = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 2)))

        If wCdCtg = "" Or wCdCd = "" Or wCdCtg <> ADC("CmCtg") Or wCdCd <> ADC("CmCd") Then _
          DispMsg "Customer code cannot be blank and needs to match the entry value. Cannot copy [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        wCdOurDmCd = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 3)))
        If wCdOurDmCd = "" Then DispMsg "Cannot Copy as our Design Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        wCdOurDmSz = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 4)))
        wCdOurSfx = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 5)))
    
        
        
        If (wCdOurDmSz <> "") And (Not moCn.RecSeek("Select PMCd From Param where  " + _
                                   "PTyp= 'DMSZ' and PMCd = '" + wCdOurDmSz + "' and PSCd = '' ")) Then _
            DispMsg "Cannot Copy as Invalid our Design Size Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                            
        If (wCdOurDmSz <> "") And (Not moCn.RecSeek("Select 'a' From DsgMst where  " + _
                                   "DmCd = '" + wCdOurDmCd + "'")) Then _
            DispMsg "Cannot Copy as Invalid our Design Code [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
      
        'select * from CustDsg where CdCtg='+wCdCtg+"' and CdCd='"+wcdcd+"' and CdOurDmCd='"+wCdOurDmCd+"' and CdOurSfx='"+wCdOurDmSfx+"'"
                
        wCdCustDmCd = UCase((ReadCell(ADC("WXLSHTNO"), i, 6)))
        wCdDesc = UCase((ReadCell(ADC("WXLSHTNO"), i, 7)))

        If Len(wCdOurDmCd) > 15 Then DispMsg "Cannot copy as our Design Code length is more than 15 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If Len(wCdOurDmSz) > 5 Then DispMsg "Cannot copy as our Design Size length is more than 5 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If Len(wCdOurSfx) > 15 Then DispMsg "Cannot copy as our Design Suffix length is more than 15 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub  '5717 sfx to 15 chr
        If Len(wCdCustDmCd) > 40 Then DispMsg "Cannot copy as Customer Design Code length is more than 40 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If Len(wCdDesc) > 120 Then DispMsg "Cannot copy as description is more than 120 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            
        If wCdCustDmCd = "" Then DispMsg "Cannot Copy as Customer Design Code cannot be blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
            If moCn.RecSeek("Select 'a' from CustDsg where CdCtg='" + ADC("CmCtg") + "' and CdCd='" + ADC("CmCd") + "' and CdOurDmCd= '" + wCdOurDmCd + "' and CdOurDmSz='" + wCdOurDmSz + "' and CdOurSfx='" + wCdOurSfx + "'") Then
            DispMsg "Cannot copy as this combination already exists. [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            End If
            If moCn.RecSeek("Select 'a' from #TmpXls where CdCtg='" + ADC("CmCtg") + "' and CdCd='" + ADC("CmCd") + "' and CdOurDmCd= '" + wCdOurDmCd + "' and CdOurDmSz='" + wCdOurDmSz + "' and CdOurSfx='" + wCdOurSfx + "'") Then
            DispMsg "Cannot copy as this combination is defined already in the Excel file. [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            End If
        
        moCn.Execute ("Insert Into #TmpXls Values (" + "'" + wCdCtg + "', '" + wCdCd + "', '" + wCdOurDmCd + "'," + _
                        "'" + wCdCustDmCd + "', '" + wCdDesc + "', '" + wCdOurDmSz + "', '" + _
                        wCdOurSfx + "')")
                        
         
    Next

    CloseExcel
    
    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls ")
'    IF WO_RSEXCELwo_rsExcel.MoveFirst
    
    Dim wRow As Double
    
    If wo_rsExcel.RecCount > 0 Then
      With GrdCustDsg
        .StartCopy
        MWLib.BeginProcess Me, "Copying Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
                  
          .Value(wRow, "CdOurDmCd") = wo_rsExcel!CdOurDmCd
          .Value(wRow, "CdOurSfx") = wo_rsExcel!CdOurSfx
          .Value(wRow, "CdOurDmSz") = wo_rsExcel!CdOurDmSz
          .Value(wRow, "CdCustDmCd") = wo_rsExcel!CdCustDmCd
          .Value(wRow, "CdDesc") = wo_rsExcel!CdDesc

          If Not .SaveRec(wRow) Then
            DispMsg "Error trying to save " + wo_rsExcel!CdOurDmCd + ":" + wo_rsExcel!CdOurSfx + ":" + wo_rsExcel!CdOurDmSz, etError
          End If
          
          wo_rsExcel.MoveNext
        Loop
        
        MWLib.EndProcess Me
        .EndCopy
      End With
    End If

Exit Sub

ErrHdlr:
  If Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
  ElseIf Err.Number = 70 Then
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If

End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable all the command buttons except CmdCpy
  '*** If CmCtg= 'C' then
      'Set the default value of CmMulBy as 1, CmFixPrc as hCmFixPrc from Head
      'also set the default values for the fields CmFixPrc, CmGldAs, CmLabAs, CmIWtEqOrd, CmIWtFrOrd,
      'CmIGldRtEqOrd, CmDiaRtFA, CmCSRtFA from the Head file
  '*** Disable all the tab frames except the frame of the current tab of the Tab Page
'  Dim ws_CoCd() As String, i As Integer, ws_ValidUsrForEdit As Boolean
'  ws_ValidUsrForEdit = True
'  If adc.AddFndMode = xFndMode And ws_MultiCoMod And (adc("CmCtg") = "C" Or adc("CmCtg") = "S" Or adc("CmCtg") = "P") Then
'        ws_CoCd = Split(adc("CmValidCoCd"), ",")
'        If IsArray(ws_CoCd) Then
'          For i = 0 To UBound(ws_CoCd)
'            Cancel = Not moCn.RecSeek("Select 1 From Head where " + _
'                                      "HCd = 'ZSELF' and HCoCd = '" + Replace$(Replace$(ws_CoCd(i), "'", "''"), """", """""") + "'")
'            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Company Code": Exit Sub
'
'            If moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + adc.MenuCd + "' And " + _
'                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + ws_CoCd(i) + "'+',%')") Then
'                ws_ValidUsrForEdit = True
'                Exit For
'            Else
'                ws_ValidUsrForEdit = False
'            End If
'                        'If ((ADC("FgFrBLoc") <> "" And InStr(1, "," + UCase(gs_ValidPrdLocSeq) + ",", "," + UCase(ADC("FgFrBLoc")) + ",") <= 0)
 '         Next i
'          If Not ws_ValidUsrForEdit Then
'                Cancel = True
'                ErrMsg = "User Doesn't Have Access To This Customer/Supplier/PL Code"
''                ADC.AllowClose = True
'                adc.AllowDelete = False
'                adc.AllowSave = True
'                Exit Sub
'            End If
'        End If
'  End If
'    If ADC("OmPrtKey") = ctPrevPrtn Then Cancel = True: ErrMsg = "Cannot Edit Records From Previous Partition ": Exit Sub
'

  Call EnaDisaCmds(False)
  CmdCpy.Enabled = False
  CmdCpyRm.Enabled = False  'Sachin - 4.1.0.0
  CmdImpExc.Enabled = False
  '*** Alternate Code
  '  FraNKeyAll.Enabled = True
  '  FraCpy.Enabled = False
  albl_BuyAdd.Caption = "Buyer's Address"
  If ADC("CmCtg") = "C" Or ADC("CmCtg") = "T" Or ADC("CmCtg") = "S" Or ADC("CmCtg") = "P" Then
    If ADC("CmMulby") = 0 Then ADC("CmMulby") = 1
    If ADC("CmMrpMulby") = 0 Then ADC("CmMrpMulby") = 1     '4.1.4
    If ADC.AddFndMode = xaddmode Then
      Dim wRsHead As MwfLib.MDORowSet
      
      ADC("CmTCSYN") = "N"
      ADC("CMAllwNewSOYN") = "Y"
      ADC("CmInvFtRem") = moCn.GetFldVal("Select hFtRem From Head Where HCoCd='" + ctSelfCoCd + "' And HCd='" + ctSelfCmCd + "'")   '4.1.2
      ADC("CmInvShowGrsWtYN") = "N"     '4.1.2
      ADC("CmInvShowDiaWtYN") = "N"     '4.1.2
      
      ADC("CmSeekZSelfRt") = "Y"        '4.1.4
      ADC("CmSeekZSelfPDCM") = "Y"      '4.1.4
      ADC("CmCreateDt") = moCn.SrvrDate   '6.1B-24
      Set wRsHead = moCn.OpenRes("Select * From Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
      If Not (wRsHead.EOF Or wRsHead.BOF) Then
        ADC("CmFixPrc") = wRsHead!hCmFixPrc
        ADC("CmGldAs") = wRsHead!hCmGldAs
        ADC("CmLabAs") = wRsHead!hCmLabAs       '****** Sachin 2.13.0 - 10-05-2006
        ADC("CmIWtEqOrd") = wRsHead!hCmIWtEqOrd
        ADC("CmIWtFrOrd") = wRsHead!hCmIWtFrOrd
        ADC("CmIGldRtEqOrd") = wRsHead!hCmIGldRtEqOrd
        
        '****** Zubin 3.02.01 - 26/03/08
        ADC("CmILabWtFrOrd") = wRsHead!hCmILabWtFrOrd
        '****** Zubin 3.02.01 - 26/03/08
        'EmrSv - Uma
        ADC("CmCell") = wRsHead!HCELL
        ADC("CmDiaRtFA") = wRsHead!hCmDiaRtFA
        ADC("CmCSRtFA") = wRsHead!hCmCSRtFA
        ADC("CmHInsBy") = wRsHead!HInsBy
        ADC("CmHRndOrdSalPrc") = wRsHead!HRndOrdSalPrc
        '***************Geeta****************Emr207
        ADC("CmVessel") = wRsHead!HVessel
        '***************Geeta****************Emr207
        ' Zubin 213 (Bef 213)
        'ADC("CmCurCd") = moCn.GetFldVal("Select PmCd from Param where PTyp= 'CURNCY' and PmCd = '" + ctDefCurncy + "'")
        ' Zubin 213
        ADC("CmCurCd") = moCn.GetFldVal("Select HBaseCurCd from Head where HCoCd= '" + gs_CoCd + "' and " + _
                          "HCd= '" + ctSelfCmCd + "' ")
        If ADC("CmCtg") = "C" Then
          ADC("CmVatYn") = "N"
          ADC("CmCstYn") = "N"
          ADC("CmOctroiYn") = "N"
        End If
        'vk.24 getting value from head for new customer
        ADC("CmInvByBagYN") = wRsHead!hInvByBagYN
      End If
    End If
  ElseIf ADC("CmCtg") = "Z" Then
    albl_BuyAdd.Caption = "Register Office"
    If ADC.AddFndMode = xaddmode Then _
    ADC("CmName") = moCn.GetFldVal("Select HName from Head where HCd = '" + ctSelfCmCd + "' and " + _
                              "HCoCd = '" + ADC("CmCd") + "' ")
  End If
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  If ADC.AddFndMode = xaddmode Then ADC("CmValidYN") = "Y"
' ###########################################  Manoj  ###########################################
  Dim j As Single
  For j = 0 To TabCustDet.Tabs - 1
    FraTabCustDet(j).Enabled = False
  Next j
  FraTabCustDet(TabCustDet.Tab).Enabled = True
  Set wRsHead = Nothing
  
  If ADC("CmInvAllBags") = "" Then ADC("CmInvAllBags") = "N"
  '*** jay 2.14 *** [01/03/07]
  'If GetUsrLevel > 5# Then ADC.AllowDelete = False
  If ms_UsrLvl > 5# Then ADC.AllowDelete = False
  '*** jay 2.14 *** [01/03/07]
  CmdSC.Enabled = False   '6.1-90 scope will not enable after click of find

End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** Jay 2.14 ***
  If (UCase(IdName) = UCase("CmName") Or UCase(IdName) = UCase("CmValidYN")) Then
    '*** jay 2.14 *** [01/03/07]
    If ms_UsrLvl > 5# Then Cancel = True: ErrMsg = "Access To Edit This Field Is Denied To User " + gs_UsrCd: Exit Sub
    'If GetUsrLevel > 5# Then Cancel = True: ErrMsg = "Access To Edit This Field Is Denied To User " + gs_UsrCd: Exit Sub
    '*** jay 2.14 *** [01/03/07]
  End If
  '*** Jay 2.14 ***
  
  Select Case UCase(IdName)
  Case Is = UCase("CmCd")
    Call HlpList.CustCd(ADC("CmCtg"), False)
    '*** Jay 2.14 ***
  Case Is = UCase("CmValidCoCd")
    HlpList.MultiSelect = True
    Call HlpList.hCoCd

' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("CMAllwNewSOYN")
      Call HlpList.PMCd("YN")
  Case Is = UCase("CMAllwNewSODESC")
      If ADC("CMAllwNewSOYN") = "Y" Then
        Cancel = True
        ErrMsg = "Description not required when allow SO YN is Yes"
      End If
  Case Is = UCase("CmValidYN")
    Call HlpList.PMCd("YN")
' ###########################################  Manoj  ###########################################
  Case Is = UCase("CmGldAs")
    Call HlpList.PMCd("GLDAS")
  
  '****** Sachin 2.13.0 - 10-05-2006
  Case Is = UCase("CmLabAs")
    Call HlpList.PMCd("GLDAS")
  '****** Sachin 2.13.0 - 10-05-2006
    
  Case Is = UCase("CmIWtEqOrd")
    Call HlpList.PMCd("YN")
  'UMA **** ADD
  Case Is = UCase("CMCELL")
    Call HlpList.PMCd("CELL")
    
  '*** When CmIWtEqOrd= 'Y' then CmIWtFrOrd is made blank in recalc, and read only
  Case Is = UCase("CmIWtFrOrd")
    If ADC("CmIWtEqOrd") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    Call HlpList.PMCd("IWTFRORD")
  Case Is = UCase("CmIGldRtEqOrd")
    Call HlpList.PMCd("YN")
  Case Is = UCase("CmDiaRtFA")
    Call HlpList.PMCd("FA")
  Case Is = UCase("CmCSRtFA")
    Call HlpList.PMCd("FA")
  Case Is = UCase("CmFixPrc")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wCustFr")    '4.1.0.0    '*** used in the frame FraCpy.
    Call HlpList.CustCd("C")
  Case Is = UCase("wCustrmFr")  '578.19
    Call HlpList.CustCd(ADC("CmCtg"))
  Case Is = UCase("CmLkUpRmRt")
    Call HlpList.CustCd("P")
  Case Is = UCase("CmLkUpLabRt")
    Call HlpList.CustCd("P")
  Case Is = UCase("CmLkUpMetLs")
    Call HlpList.CustCd("P")
  'pg.5- calling customer of cmctg=P
  Case Is = UCase("CmLkUpMrpDisc")
    Call HlpList.CustCd("P")
  Case Is = UCase("CmRegnCd")
    Call HlpList.PMCd("REGION")
  Case Is = UCase("CmHRndOrdSalPrc")
    Call HlpList.PMCd("RNDOPT")
  Case Is = UCase("CmSalPer")
    Call HlpList.PMCd("SP")
  Case Is = UCase("CmCurCd")
    Call HlpList.PMCd("CURNCY")
'****** Shilpa *********
  '*** Jay 3.1.1 ***  [CmOurBkCd]
  Case Is = UCase("CmBkCd"), UCase("CmOurBkCd")
    Call HlpList.BkCd("B")
  Case Is = UCase("CmAgCd")
    Call HlpList.BkCd("A")
'****** Shilpa *********
  '*** Jay  2.14(LclInv) ***
  Case Is = UCase("CmBillCo")
    Call HlpList.hCoCd
  Case Is = UCase("CmVatYn"), UCase("CmCstYn"), UCase("CmOctroiYn")
    Call HlpList.PMCd("YN")
  '*** Jay  2.14(LclInv) ***
  
  '*** Jay 3.2.0 [Stmp Instr]
  Case Is = UCase("WCMSTMPCD")
      Call HlpList.PMCd("STMPINST")
  '*** Jay 3.2.0 [Stmp Instr]
  
  '****** Zubin 3.02.01 - 26/03/08
  Case Is = UCase("CmILabWtFrOrd"), UCase("CMINVSHOWGRSWTYN"), UCase("CMINVSHOWDIAWTYN"), UCase("CmInvAllBAgs")     ', UCase("CmInvShowNetWtYN"), UCase("CmInvShowGrsWtYN"), UCase("CmInvShowDiaWtYN"), UCase("CmInvShowCsWtYN")       'Addl. fields in 4.1.2
    Call HlpList.PMCd("YN")
  '****** Zubin 3.02.01 - 26/03/08
  
  Case Is = UCase("wFrRmCtg"), UCase("wToRmCtg")
    Call HlpList.PMCd("RMCTG")
  
 'Sachin 4.1.0.0
  Case Is = UCase("CmGrp")
    Call HlpList.PMCd("CMGRP")
  
  Case Is = UCase("CmTyp")
    Call HlpList.PMCd("CMTYP")
    
  Case Is = UCase("CmPod")
    Call HlpList.PMCd("PORTCD")
  
  Case Is = UCase("CmFinDstCd")
    Call HlpList.PMCd("CONTRYCD")
  'Sachin 4.1.0.0
  
  Case Is = UCase("CmSeekZSelfRt"), UCase("CmSeekZSelfPDCM")        '4.1.4
    Call HlpList.PMCd("YN")
  'PYMTRM-Help for PayTermCode
  Case Is = UCase("CmPayTermsCd")
    Call HlpList.PMCd("PYMTRM")
  'vk.24 calling help for invoice by bag
  Case Is = UCase("CmInvByBagYN")
    Call HlpList.PMCd("YN")
  'vm.3 calling help for Apply Multiply Factor to Invoice.
  'if Multiply BY value is 1 then cannot change this value, help will not work
  Case Is = UCase("CmOrdMulByYN")
    If ADC("CmMulBy") = 1 Then Cancel = True: Exit Sub
    Call HlpList.PMCd("YN")
  'sv.44 calling help for PDCMLkUp & apply TCS Y/N
  Case Is = UCase("CmLkUpPDCM")
    Call HlpList.CustCd("C")
  Case Is = UCase("CmTCSYN")
    Call HlpList.PMCd("YN")
  '6.1-90 help for scope fields
  Case Is = UCase("wScDmCtgFr"), UCase("wScDmCtgTo")
    Call HlpList.PMCd("DMCTG")
  Case Is = UCase("wScDmCdFr"), UCase("wScDmCdTo")
    Call HlpList.DmCd("DM")
    
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** CmIWtFrOrd on CmIWtEqOrd         Normal
  'If CmIWtEqOrd= 'Y' then ''
      
  Select Case UCase(IdName)
  Case Is = UCase("CmIWtFrOrd")
    If ADC.Mode = xNorm Then
      If ADC("CmIWtEqOrd") = "Y" Then ADC("CmIWtFrOrd") = ""
    End If
    
  '*** Jay 3.2.0 [Stmp Inst]
  Case Is = UCase("CMSTMPINST")
    If ADC.Mode = xNorm Then
      If ADC("WCMSTMPCD") <> "" Then ADC("CMSTMPINST") = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'STMPINST' and PmCd= '" + ADC("WCMSTMPCD") + "' and PSCd= ''")
    End If
  '*** Jay 3.2.0 [Stmp Inst]
  
  'Sachin 4.1.0.0
  Case Is = UCase("CMPORTOFDISCH")
    If ADC.Mode = xNorm Then
      If ADC("CMPOD") <> "" Then ADC("CMPORTOFDISCH") = moCn.GetFldVal("Select PDesc from Param where PTyp = 'PORTCD' and PmCd= '" + ADC("CMPOD") + "' and PSCd= ''")
    End If

  Case Is = UCase("CMFINDEST")
    If ADC.Mode = xNorm Then
      If ADC("CMFINDSTCD") <> "" Then ADC("CMFINDEST") = moCn.GetFldVal("Select PDesc from Param where PTyp = 'CONTRYCD' and PmCd= '" + ADC("CMFINDSTCD") + "' and PSCd= ''")
    End If
  'Sachin 4.1.0.0
  'PYMTRM-If PayTerm description field is blank then it gets description from PayTermcode
  'else it remains as such
  Case Is = UCase("DMPAYTERMS")
    If ADC.Mode = xNorm Then
      If ADC("DMPAYTERMS") = "" And ADC("CMPAYTERMSCD") <> "" Then ADC("DMPAYTERMS") = moCn.GetFldVal("Select PDesc225 from Param where PTyp = 'PYMTRM' and PmCd= '" + ADC("CMPAYTERMSCD") + "' and PSCd= ''")
    End If
  'vm.3 if MultiplyBy value is 1.00, then recalc its value to 'N' else gets value from Head
  Case Is = UCase("CmOrdMulByYN")
    If ADC.Mode = xNorm Then
      If ADC("CmMulBy") = 1 Then
        ADC("CmOrdMulByYN") = "N"
      Else
        ADC("CmOrdMulByYN") = moCn.GetFldVal("Select hOrdMulByYN from Head where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
      End If
    End If
  Case Is = UCase("CMALLWNEWSODESC")
    If ADC("CMAllwNewSOYN") = "Y" Then
      ADC("CMALLWNEWSODESC") = ""
    End If
      
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** CmCd             Not Blank
  '*** CmGldAs          valid PMCd('GLDAS')
  '*** CmLabAs          valid PMCd('GLDAS')
  '*** CmIWtEqOrd       valid PMCd('GLDAS')
  '*** CmIWtFrOrd       blank or valid PMCd('IWTFRORD')
  '*** CmIGldRtEqOrd    valid PMCd('YN')
  '*** CmDiaRtFA        valid PMCd('FA')
  '*** CmCSRtFA         valid PMCd('FA')
  '*** CmFixPrc         valid PMCd('YN')
  '*** CmMulby          > 0
  '*** wCustFr          blank or valid CmCd from CustMst(CmCtg= 'C')       (This Field is in the Alloy Copy frame)
  Dim ws_CoCd() As String, i As Integer
  'stellar.1- calling a function to check Quotes exists in given value
  Dim wChkQtsErrMsg As String
  If InStr(1, "|" + ms_ChkQtsIdName + "|", "|" + IdName + "|") > 0 Then
    wChkQtsErrMsg = ChkQuotes(Me, IdName, pv_NewValue)
    If wChkQtsErrMsg <> "" Then Cancel = True: ErrMsg = wChkQtsErrMsg: Exit Sub
  End If

  Select Case UCase(IdName)
  Case Is = UCase("CmCd")
    If Trim(pv_NewValue) = "" Then Cancel = True: ErrMsg = "Customer Code Cannot Be Blank": Exit Sub
    '*** For Ctg = "Z"
    If ADC("CmCtg") = "Z" Then
      Cancel = Not moCn.RecSeek("select HCoCd from Head where HCd = '" + ctSelfCmCd + "' and " + _
                                "HCoCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Company Code": Exit Sub
    End If
  Case Is = UCase("CmName")
    '*** For Ctg = "Z"
    If ADC("CmCtg") = "Z" Then
      '****** Sachin [Before 2.14.0 Next 12-02-07] ******
      'Cancel = Not moCn.RecSeek("select HName from Head where HCd = '" + ctSelfCmCd + "' and " + _
                           "HCoCd = '" + ADC("CmCd") + "' and HName = '" + pv_NewValue + "'")
      'ErrMsg = "Invalid Company Name": Exit Sub
      
      '****** Sachin 2.14.0 Next [12-02-07] ******
        Cancel = ADC("CmName") <> "" And Not moCn.RecSeek("Select CmName from CustMst Where CmCtg= 'Z' And CmCd = '" + ADC("CmCd") + "'" + _
                             " And CmName = '" + pv_NewValue + "'")
        If Cancel = True Then ErrMsg = "Cannot Change Company Name": Exit Sub
        '****** Sachin 2.14.0 Next [12-02-07] ******
    End If
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("CmValidCoCd")
        ws_CoCd = Split(pv_NewValue, ",")
        If IsArray(ws_CoCd) Then
          For i = 0 To UBound(ws_CoCd)
            Cancel = Not moCn.RecSeek("Select 1 From Head where " + _
                                      "HCd = 'ZSELF' and HCoCd = '" + Replace$(Replace$(ws_CoCd(i), "'", "''"), """", """""") + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Company Code": Exit Sub
            Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + ws_CoCd(i) + "'+',%')")
            If Cancel = True Then ErrMsg = "User Does Not have Access to (" + CStr(i + 1) + ") Company Code": Exit Sub
          Next i
        End If

  Case Is = UCase("CmValidYN"), UCase("CMAllwNewSOYN")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Valid (Y/N)": Exit Sub
' ###########################################  Manoj  ###########################################
  Case Is = UCase("CmRegnCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                         " PTyp= 'REGION' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Region Code": Exit Sub
  Case Is = UCase("CmBkCd")
    Cancel = (Not moCn.RecSeek("select BkCd from BnkAgMst where " + _
                         " BkCtg= 'B' and BkCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Bank Code": Exit Sub
  '*** Jay 3.1.1 ***  [CmOurBkCd]
  Case Is = UCase("CmOurBkCd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select BkCd from BnkAgMst where " + _
                         " BkCtg= 'B' and BkCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Our Bank Code": Exit Sub
  '*** Jay 3.1.1 ***  [CmOurBkCd]
  Case Is = UCase("CmAgCd")
    Cancel = (Not moCn.RecSeek("select BkCd from BnkAgMst where " + _
                         " BkCtg= 'A' and BkCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Agent Code": Exit Sub
  Case Is = UCase("CmGldAs")
      '*** (Bef 2.14)
      'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
      '                     " PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
      '*** (Bef 2.14)
      
      '*** (Jen 2.14)
      'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' " + _
      '         "and PMCd = '" + pv_NewValue + "' and not (',' + PDesc225 + ',' like '%,G,%' Or ',' + PDesc225 + ',' like '%,P,%' Or ',' + PDesc225 + ',' like '%,S,%') "))
      
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where PTyp= 'GLDAS' " + _
               "and PMCd = '" + pv_NewValue + "' "))
      '*** (Jen 2.14)
      
      ErrMsg = "Invalid Gold As Value": Exit Sub
  
  '****** Sachin 2.13.0 - 10-05-2006
  Case Is = UCase("CmLabAs")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Labour As Value": Exit Sub
  '****** Sachin 2.13.0 - 10-05-2006
  
  Case Is = UCase("CmIWtEqOrd")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invoice Wts From Order (Y/N)": Exit Sub
  Case Is = UCase("CmIWtFrOrd")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'IWTFRORD' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Wts From Order": Exit Sub
  Case Is = UCase("CmIGldRtEqOrd")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invoice Gold Rate From Order (Y/N)": Exit Sub
  Case Is = UCase("CmDiaRtFA")
      '*** (Bef 2.14)
      'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
      '                     " PTyp= 'FA' and PMCd = '" + pv_NewValue + "'"))
      '*** (Bef 2.14)
      Cancel = pv_NewValue <> "A"    '*** (Jen 2.14)
      ErrMsg = "Enter Diamond Rate Favourable/Actual": Exit Sub
  Case Is = UCase("CmCSRtFA")
      '*** (Bef 2.14)
      'Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
      '                     " PTyp= 'FA' and PMCd = '" + pv_NewValue + "'"))
      '*** (Bef 2.14)
      Cancel = pv_NewValue <> "A"    '*** (Jen 2.14)
      ErrMsg = "Enter Colour Stone Rate Favourable/Actual": Exit Sub
      
  Case Is = UCase("CmFixPrc")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Fixed Price (Y/N)": Exit Sub
      
  Case Is = UCase("CmMulby")
      If pv_NewValue = 0 And ADC("CmCtg") = "C" Then _
      Cancel = True: ErrMsg = "Multiplying Factor Should Be > 0": Exit Sub
  
  Case Is = UCase("CmMrpMulby")                         '4.4.0
      If pv_NewValue = 0 And ADC("CmCtg") = "C" Then
        Cancel = True: ErrMsg = "MRP Multiplying Factor Should Be > 0": Exit Sub
      End If
    
  Case Is = UCase("wCustFr")    'Sachin - 4.1.0.0
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'C' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Source Customer Code": Exit Sub
  '578.19
  Case Is = UCase("wCustRmFr")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = '" + ADC("CmCtg") + "' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Source Customer Code": Exit Sub
  Case Is = UCase("CmLkUpRmRt")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'P' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Price List Customer Code For Rm Rate Look Up": Exit Sub
  Case Is = UCase("CmLkUpLabRt")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'P' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Price List Customer Code For Lab Rate Look Up": Exit Sub
  Case Is = UCase("CmLkUpMetLs")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'P' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Price List Customer Code For Metal Loss Look Up": Exit Sub
  'pg.5- Checking valid "P" customer, blank allowed
  Case Is = UCase("CmLkUpMrpDisc")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'P' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Price List Customer Code For MRP Discount Look Up": Exit Sub
  Case Is = UCase("CmHRndOrdSalPrc")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'RNDOPT' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Specify Method In Which Sales Price Will Be Rounded In The Order Entry For This Customer": Exit Sub
  Case Is = UCase("CmSalPer")
    Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                                      "PTyp= 'SP' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Sales Executive Code": Exit Sub
  Case Is = UCase("CmCurCd")
    Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                     "PTyp= 'CURNCY' and PMCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Currency Code": Exit Sub
  '*** Jay  2.14(LclInv) ***
  Case Is = UCase("CmBillCo")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select 'x' from Head where HCocd= '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Company Code.": Exit Sub
  Case Is = UCase("CmVatYn")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid VAT Applicable (Yes/No) ": Exit Sub
  Case Is = UCase("CmCstYn")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid CST Applicable (Yes/No) ": Exit Sub
  Case Is = UCase("CmOctroiYn")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Octroi Applicable (Yes/No) ": Exit Sub
  '*** Jay  2.14(LclInv) ***
  '*** Jay 3.2.0 [Stmp Instr]
  Case Is = UCase("WCMSTMPCD")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'STMPINST' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Stamping Instruction Code Not Defined": Exit Sub
  '*** Jay 3.2.0 [Stmp Instr]

  '****** Zubin 3.02.01 - 26/03/08
  Case Is = UCase("CmILabWtFrOrd")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Labor Weight From Order (Y/N)": Exit Sub
  '****** Zubin 3.02.01 - 26/03/08
  Case Is = UCase("CmInvAllBags")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Inv-Wait for all Bags (Y/N)": Exit Sub
  
  'Sachin - 4.1.0.0 - Should b a valid Customer Grp and Typ
  Case Is = UCase("CmGrp")
      Cancel = (ADC("CmCtg") = "C" Or pv_NewValue <> "") And Not moCn.RecSeek("Select PMCd from Param where PTyp= 'CMGRP' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
      ErrMsg = "Invalid Customer Group": Exit Sub

  Case Is = UCase("CmTyp")
    Cancel = (ADC("CmCtg") = "C" Or pv_NewValue <> "") And Not moCn.RecSeek("Select PMCd from Param where PTyp= 'CMTYP' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Customer Type": Exit Sub
  
  Case Is = UCase("CmPod")
    Cancel = (ADC("CmCtg") = "C" Or pv_NewValue <> "") And Not moCn.RecSeek("Select PMCd from Param where PTyp= 'PORTCD' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Port of Discharge Code": Exit Sub
  
  Case Is = UCase("CmFinDstCd")
    Cancel = (ADC("CmCtg") = "C" Or pv_NewValue <> "") And Not moCn.RecSeek("Select PMCd from Param where PTyp= 'CONTRYCD' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Final Destination Code": Exit Sub
  'Sachin - 4.1.0.0
  
  'Sachin - 4.2.0.0
  Case Is = UCase("CMINVSHOWGRSWTYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Default Invoice Net Wt Option (Y/N)": Exit Sub
  
  Case Is = UCase("CMINVSHOWDIAWTYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "' and PSCd= ''")
    ErrMsg = "Invalid Default Invoice Grs Wt Option (Y/N)": Exit Sub
  'Sachin - 4.2.0.0
  
  Case Is = UCase("CmSeekZSelfRt")  '4.1.4
      Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Rate seek to ZSelf depth (Y/N)": Exit Sub
  
  Case Is = UCase("CmSeekZSelfPDCM")  '4.1.4
      Cancel = (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "PDCM seek to ZSelf depth (Y/N)": Exit Sub
  'UMA **** ADD
  Case Is = UCase("CMCELL")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'CELL' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid CELL Name": Exit Sub
  'PYMTRM-If PayTermCode is non blank then it checks valid PayTerm code, blank allowed
  Case Is = UCase("CmPayTermsCd")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'PYMTRM' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Payment Term Code": Exit Sub
  'vk.24 validating invoice by bag
  Case Is = UCase("CmInvByBagYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Invoice By Bag (Y/N)": Exit Sub
  'vm.3 validating  Apply Multiply Factor to Invoice
  Case Is = UCase("CmOrdMulByYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Apply Multiply Factor to Invoice (Y/N)": Exit Sub
  'sv.44 validating customer for PDCM lookup
  Case Is = UCase("CmLkUpPDCM")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select CmCd From CustMst Where " + _
                                    "CmCtg = 'C' and CmCd = '" + pv_NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Customer Code For PDCM Look Up": Exit Sub
  'sv.44 validating  Apply TCS
  Case Is = UCase("CmTCSYN")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Apply TCS (Y/N)": Exit Sub

  End Select
End Sub
Private Sub CmdCpyRm_Click()
    'Sachin - 4.1.0.0
    Call DispFra(CpyRm)
    CmdSC.Enabled = False   '6.1-90 disable scope button
End Sub
Private Sub CmdCpyRMGo_Click()
    'Sachin - 4.1.0.0
    Call CpyRmDet
    Call CmdCpyRm_Click
End Sub

Private Sub GrdCustDsg_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0
  If GrdCustDsg.Value(RowNum, "CdPrtKey") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Designs": Exit Sub
  
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("CdOurDmCd")
    Call HlpList.DmCd("DM")
  Case Is = UCase("CdOurSfx")
    'Call HlpList.OdSfx(ADC("CmCd"), GrdCustDsg.Value(RowNum, "CdOurDmCd"))
  Case Is = UCase("CdOurDmSz")
    Call HlpList.PMCd("DMSZ")
  End Select
End Sub
Private Sub GrdCustDsg_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** CdOurDmCd           valid DmCd from DsgMst(DmTcTyp= 'DM')
  '*** CdCustDmCd          Not Blank
  
  With GrdCustDsg
    Select Case ColName
    Case Is = UCase("CdOurDmCd")
      Cancel = Not moCn.RecSeek("Select DmCd From DsgMst Where " + _
                           "DmTcTyp= 'DM' And DmCd = '" + NewValue + "' And DmPrtKey='" + ctCurrPrtn + "'")
      ErrMsg = "Invalid Design Code": Exit Sub
    Case Is = UCase("CdCustDmCd")
      If NewValue = "" Then Cancel = True: ErrMsg = "Invalid Customer Design": Exit Sub
    Case Is = UCase("CdOurDmSz")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                                    "PTyp= 'DMSZ' And PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Design Size": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdCustDsg_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdCustDsg
    .Store "CdCtg", ADC("CmCtg")
    .Store "CdCd", ADC("CmCd")
  End With
End Sub

Private Sub GrdCustDsg_RowWhen(ByVal RowNum As Integer)

'****** Sachin 3.02
If GrdCustDsg.IsNew(RowNum) Then GrdCustDsg.Value(RowNum, "CdPrtKey") = ctCurrPrtn

End Sub

Private Sub GrdCustInst_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdCustInst
    Select Case ColName
    Case Is = UCase("CiDmCtg")
                Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'DMCTG' and PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Design Category": Exit Sub
    Case Is = UCase("CiCustPrdInst")
                Cancel = NewValue = ""
      ErrMsg = "Customer Production Instruction Cannot Be Blank": Exit Sub
    End Select
  End With

End Sub

Private Sub GrdCustInst_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
Select Case ColName
  Case Is = UCase("CiDmCtg")
    Call HlpList.PMCd("DMCTG")
  End Select
End Sub

Private Sub GrdCustInst_InitKey(ByVal RowNum As Integer)
  With GrdCustInst
      .Store "CiCtg", ADC("CmCtg")
      .Store "CiCd", ADC("CmCd")
  End With

End Sub

Private Sub GrdCustInst_RowWhen(ByVal RowNum As Integer)
    If GrdCustInst.IsNew(RowNum) Then GrdCustInst.Value(RowNum, "CiPrtKey") = ctCurrPrtn
End Sub

Private Sub GrdCustInst_SetRecSource()
  GrdCustInst.RecSource = "Select * from CustInst " + _
                         "Where CiCtg ='" + ADC("CmCtg") + "' " + _
                         "And CiCd = '" + ADC("CmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And CiPrtKey='" + ctCurrPrtn + "' ", "") + " Order By CiDmCtg"

End Sub

Private Sub GrdCustRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0
  If GrdCustRm.Value(RowNum, "CrPrtKey") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Designs": Exit Sub
  
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("CrRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("CrOurRmCd")
    Call HlpList.RmCd("N", "'" + GrdCustRm.Value(RowNum, "CRRMCTG") + "'")
  Case Is = UCase("CrFrLn"), UCase("CrToLn")
    If GrdCustRm.Value(RowNum, "CrRmCtg") = "D" Or GrdCustRm.Value(RowNum, "CrRmCtg") = "C" Then
       Cancel = False
    Else
      Cancel = True
    End If
  End Select
End Sub
Private Sub GrdCustRm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** CdOurDmCd           valid DmCd from DsgMst(DmTcTyp= 'DM')
  '*** CdCustDmCd          Not Blank
  
  With GrdCustRm
    Select Case ColName
    Case Is = UCase("CrRmCtg")
      Cancel = Not moCn.RecSeek("Select PmCd From Param Where " + _
                           "PTyp= 'RMCTG' And PmCd = '" + NewValue + "'")
      ErrMsg = "Invalid Rm Category": Exit Sub
    Case Is = UCase("CrOurRmCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where " + _
                                    "RmCtg= '" + GrdCustRm.Value(RowNum, "CrRmCtg") + "' and " + _
                                    "RmCd = '" + NewValue + "' And RmPrtKey='" + ctCurrPrtn + "'"))
      If Cancel = True Then ErrMsg = "Invalid Rm Code": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdCustRm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdCustRm
    .Store "CrCtg", ADC("CmCtg")
    .Store "CrCd", ADC("CmCd")
  End With
End Sub
Private Sub GrdCustAly_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  Select Case ColName
  Case Is = UCase("CaRmCd")
    ' Sachin 2.12
    ' ***** Manali 3.03 - 18/06/08 - 'L' added
    Call HlpList.RmCd("N", "'G','P','S', 'L'")
  Case Is = UCase("CaDmCol")
    Call HlpList.PMCd("DMCOL")
  Case Is = UCase("CaAlyCd")
    Call HlpList.RmCd("N", "'A'")
  End Select
End Sub
Private Sub GrdCustAly_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** CaRmCd           valid RmCd from RmMst(RmCtg= 'G' or 'P')
  '*** CaDmCol          valid PMCd('DMCOL')
  '*** CaAlyCd          valid RmCd from RmMst(RmCtg= 'A')
  
  With GrdCustAly
    Select Case ColName
    Case Is = UCase("CaRmCd")
      ' Sachin 2.12
      ' ***** Manali 3.03 - 18/06/08 - 'L' added
      Cancel = Not moCn.RecSeek("Select RmCd From RmMst Where " + _
                           "RmCtg in ('G','P','S', 'L') And RmCd = '" + NewValue + "' And RmPrtKey='" + ctCurrPrtn + "'")
      ErrMsg = "Invalid Raw Material Code": Exit Sub
    Case Is = UCase("CaDmCol")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp ='DMCOL' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Design Colour": Exit Sub
    Case Is = UCase("CaAlyCd")
      Cancel = Not moCn.RecSeek("Select RmCd From RmMst Where " + _
                           "RmCtg ='A' And RmCd = '" + NewValue + "' And RmPrtKey='" + ctCurrPrtn + "'")
      ErrMsg = "Invalid Alloy Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdCustAly_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdCustAly
      .Store "CaCtg", ADC("CmCtg")
      .Store "CaCd", ADC("CmCd")
  End With
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Set the current tab of the tabpage as 0
  '*** jay 2.14 *** [01/03/07]
  If ms_UsrLvl > 5# Then TabCustDet.Tab = 3: Exit Sub
  'If GetUsrLevel > 5# Then TabCustDet.Tab = 3: Exit Sub
  '*** jay 2.14 *** [01/03/07]
  ' Emr 212
  If ADC.Mode = xNorm And ADC("CmVatYn") = "Y" And ADC("CmCstYn") = "Y" Then _
    pr_Cancel = True: pr_ErrMsg = "VAT & CST Both Can Not Be Kept Yes For Customer."
  If UCase(ADC.MenuCd) <> UCase("CtlgMst") Then TabCustDet.Tab = 0
  'stellar.1- calling a function to check Quotes exists in the list of text boxes
  Dim wChkQtsErrMsg As String
  wChkQtsErrMsg = ChkQuotes(Me, ms_ChkQtsIdName, "")
  If wChkQtsErrMsg <> "" Then pr_Cancel = True: pr_ErrMsg = wChkQtsErrMsg
  If ADC("CMAllwNewSOYN") = "N" And ADC("CMAllwNewSODESC") = "" Then
        pr_ErrMsg = "Allow New SO Description needs to be filled when Allow New SO is set to No"
        pr_Cancel = True
        Exit Sub
  End If
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC  (Base Table is CustMst)
  ADC.RecSource = " Select * from CustMst " + _
                  " Where CmCtg ='" + ADC("CmCtg") + "'" + _
                  " And CmCd = '" + ADC("CmCd") + "'"
End Sub

Private Sub GrdCustDsg_SetRecSource()
  '*** Set the RecSource for GrdCustDsg (Base Table is CustDsg)
  '6.1-90 where conditions for scope fields
  Dim wCustDsgCnd As String, wCustDsgRec As MDORowSet, wsql As String, Cancel As Boolean
  ADC("wScDmCdFr").CmpStr = "DmCd>= "
  ADC("wScDmCdTo").CmpStr = "DmCd<= "
  ADC("wScDmCtgFr").CmpStr = "DmCtg>= "
  ADC("wScDmCtgTo").CmpStr = "DmCtg<= "
  
  wCustDsgCnd = ADC.RepCond
  wCustDsgCnd = IIF(wCustDsgCnd <> "", " and " + wCustDsgCnd, "")
  wCustDsgCnd = IIF(wCustDsgCnd = "", "", " and exists(Select 'x' from DsgMst where DmCd= CdOurDmCd " + wCustDsgCnd + ")")
  '6.1-90 wCustDsgCnd added. if no. of records >32760 then raise message else continue
  wsql = "Select * from CustDsg " + _
         "Where CdCtg ='" + ADC("CmCtg") + "' " + _
         "And CdCd = '" + ADC("CmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And CdPrtKey='" + ctCurrPrtn + "' ", "") + _
         wCustDsgCnd + " Order By CdOurDmCd, CdOurSfx, CdOurDmSz"

  Set wCustDsgRec = moCn.OpenRes(wsql)
  If wCustDsgRec.RecCount > 32760 Then
    Cancel = True: MsgBox "More Than 32760 Records to Fetch in Cust Design, Please Reduce the scope"
    Exit Sub
  End If
    
  GrdCustDsg.RecSource = wsql
  
  '6.1-90 where conditions for scope fields
  ADC("wScDmCdFr").CmpStr = ""
  ADC("wScDmCdTo").CmpStr = ""
  ADC("wScDmCtgFr").CmpStr = ""
  ADC("wScDmCtgTo").CmpStr = ""

End Sub

Private Sub GrdCustRm_RowWhen(ByVal RowNum As Integer)

'****** Sachin 3.02
If GrdCustRm.IsNew(RowNum) Then GrdCustRm.Value(RowNum, "CrPrtKey") = ctCurrPrtn

End Sub

Private Sub GrdCustRm_SetRecSource()
  '*** Set the RecSource for GrdCustRm (Base Table is CustRm)
  GrdCustRm.RecSource = "Select * from CustRm " + _
                         "Where CrCtg ='" + ADC("CmCtg") + "' " + _
                         "And CrCd = '" + ADC("CmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And CrPrtKey='" + ctCurrPrtn + "' ", "") + " Order By CrRmCtg, CrOurRmCd, CrFrLn, CrToLn"
End Sub

Private Sub GrdCustAly_SetRecSource()
  '*** Set the RecSource for GrdCustAly (Base Table is CustAly)
  GrdCustAly.RecSource = "Select * from CustAly " + _
                         "Where CaCtg ='" + ADC("CmCtg") + "' " + _
                         "And CaCd = '" + ADC("CmCd") + "' Order By CaRmCd, CaDmCol"
End Sub
'578-20 validating store code, store name
Private Sub GrdCustStore_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
With GrdCustStore
  Select Case ColName
  Case Is = UCase("CstCd")
      Cancel = NewValue = ""
      ErrMsg = "Store Code Cannot Be Blank": Exit Sub
  Case Is = UCase("CstName")
      Cancel = NewValue = ""
      ErrMsg = "Store Name Cannot Be Blank": Exit Sub
  End Select
End With
End Sub
'578-20 checking partition
Private Sub GrdCustStore_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  If GrdCustStore.Value(RowNum, "CstPrtKey") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Store Details": Exit Sub
     
End Sub

'578-20 initialize custstore table
Private Sub GrdCustStore_InitKey(ByVal RowNum As Integer)
'*** Initialise Key Fields
With GrdCustStore
  .Store "CstCmCtg", ADC("CmCtg")
  .Store "CstCmCd", ADC("CmCd")
End With

End Sub
'578-20 once address 1,2,3 in footer changed then updating to grid
Private Sub GrdCustStore_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'With GrdCustStore
'  Select Case UCase(ColName)
'    Case Is = UCase("CstAdd1")
'      If .Mode = fgmnorm Then
'        If .Row > 0 Then
'          If .Value(.Row, "CstAdd1") <> ADC("wCstAdd1") Then
'            .Value(.Row, "CstAdd1") = ADC("wCstAdd1")
'            .SaveRec .Row
'          End If
'        End If
'      End If
'    Case Is = UCase("CstAdd2")
'      If .Mode = fgmnorm Then
'        If .Row > 0 Then
'          If .Value(.Row, "CstAdd2") <> ADC("wCstAdd2") Then
'            .Value(.Row, "CstAdd2") = ADC("wCstAdd2")
'            .SaveRec .Row
'          End If
'        End If
'      End If
'    Case Is = UCase("CstAdd3")
'      If .Mode = fgmnorm Then
'        If .Row > 0 Then
'          If .Value(.Row, "CstAdd3") <> ADC("wCstAdd3") Then
'            .Value(.Row, "CstAdd3") = ADC("wCstAdd3")
'            .SaveRec .Row
'          End If
'        End If
'      End If
'  End Select
'End With
End Sub

'578-20 when a row selected then address details will show in footer
Private Sub GrdCustStore_RowWhen(ByVal RowNum As Integer)
With GrdCustStore
  If .IsNew(RowNum) Then .Value(RowNum, "CstPrtKey") = ctCurrPrtn
  ADC("wCstAdd1") = .Value(RowNum, "CstAdd1")
  ADC("wCstAdd2") = .Value(RowNum, "CstAdd2")
  ADC("wCstAdd3") = .Value(RowNum, "CstAdd3")
        
End With
End Sub

'578-20 Set the RecSource for GrdCustStore
Private Sub GrdCustStore_SetRecSource()
  GrdCustStore.RecSource = "Select * from CustStore " + _
                         "Where CstCmCtg ='" + ADC("CmCtg") + "' " + _
                         "And CstCmCd = '" + ADC("CmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And CstPrtKey='" + ctCurrPrtn + "' ", "") + " Order By CstCd "
                                    
End Sub
'578-20
Private Sub GrdCustStore_Validate(Cancel As Boolean)
  Cancel = GrdCustStore.Validate
End Sub

Private Sub MWCTL_BTN11_Click()

End Sub

Private Sub TabCustDet_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  '*** Enable the CmdCpy Button only when the tab= 4 (The Alloy Details Tab)
  '*** For this to work the index of the frame and the Tab Number should be the same

  FraTabCustDet(TabCustDet.Tab).Enabled = True
  FraTabCustDet(PreviousTab).Enabled = False
  If TabCustDet.Tab = 4 Then
    CmdCpy.Enabled = True
  Else
    CmdCpy.Enabled = False
  End If
  CmdCpyRm.Enabled = IIF(TabCustDet.Tab = 5, True, False)   'Sachin - 4.1.0.0
  CmdImpExc.Enabled = IIF(TabCustDet.Tab = 3 Or TabCustDet.Tab = 5, True, False)  '578.19
  '578-20 when cust store tab clicked then footer address details will display
  If TabCustDet.Tab = 8 Then
    FraAddr.Visible = True
    FraAddr.ZOrder
    If GrdCustStore.Rows > 1 Then GrdCustStore.Row = 1
  Else
    FraAddr.Visible = False
  End If
    
End Sub

Private Sub CmdCpy_Click()
  '*** Display or hide the Copy frame depending on the toggle satate
  Call DispFra(Cpy)
  CmdSC.Enabled = False
  '*** Alternative Code
  '  FraCpy.Enabled = True
  '  FraNKeyAll.Enabled = False
  '  CmdCpy.Enabled = False
  '  FraCpy.ZOrder (0)
End Sub
Private Sub CmdCpyGo_Click()
  '*** Call the Copy Routine and Close the Copy Frame
  Call CpyAlyDet
  Call CmdCpy_Click
  
  '*** Alternative Code
  '  Call CpyAlyDet
  '  FraNKeyAll.Enabled = True
  '  GrdCustAly.COL = 1
  '  GrdCustAly.SetFocus
  '  FraCpy.Enabled = False
  '  CmdCpy.Enabled = True
  '  FraNKeyAll.ZOrder (0)
End Sub
Private Sub CpyAlyDet()
  '*** Copy Option is used to copy alloy details from the specified Customer to the current customer
  '*** No rows are copied if alloy records already exist

  Dim wRow As Single
  '--- Checking if copy is possible
  If Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= 'C' " + _
     "and CmCd = '" + ADC("wCustFr") + "'") Then _
     DispMsg "Enter Proper Source Customer Code", etError: Exit Sub
  
  If GrdCustAly.Rows - 1 > 0 Then DispMsg "Cannot Copy As Alloy Details Already Exist", etError: Exit Sub
  
  '*** Beginning the copy process
  GrdCustAly.StartCopy
  MWLib.BeginProcess Me, "Copying Alloy Details ..."
  Dim wRsCustAly As MwfLib.MDORowSet
  Set wRsCustAly = moCn.OpenRes("Select * from CustAly where CaCtg= 'C' and " + _
                 "CaCd= '" + ADC("wCustFr") + "' Order By CaRmCd, CaDmCol, CaAlyCd")
  With wRsCustAly
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Alloy Details"
      GrdCustAly.AddItem
      wRow = GrdCustAly.Rows - 1
      GrdCustAly.Value(wRow, "CaRmCd") = !CaRmCd
      GrdCustAly.Value(wRow, "CaDmCol") = !CaDmCol
      GrdCustAly.Value(wRow, "CaAlyCd") = !CaAlyCd
      GrdCustAly.SaveRec (wRow)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdCustAly.EndCopy
  Set wRsCustAly = Nothing
  '*** End of the copy process
End Sub
Private Sub CpyRmDet()
  'Sachin - 4.1.0.0
  Dim wRow As Single
  Dim wCnd As String
  '578.19 CmCtg= 'S' added
  If Not moCn.RecSeek("Select CmCd From CustMst where (CmCtg= 'C' or CmCtg= 'S') " + _
     "and CmCd = '" + ADC("wCustRmFr") + "'") Then _
     DispMsg "Enter Proper Source Customer Code", etError: Exit Sub
  
  'If GrdCustRm.Rows - 1 > 0 Then DispMsg "Cannot Copy As RM Details Already Exist", etError: Exit Sub
  
  wCnd = IIF(ADC("wFrRmCtg") <> "", " And CrRmCtg>= '" + ADC("wFrRmCtg") + "'", "") + _
         IIF(ADC("wToRmCtg") <> "", " And CrRmCtg<= '" + ADC("wToRmCtg") + "'", "")
  
  '*** Beginning the copy process
  GrdCustRm.StartCopy
  MWLib.BeginProcess Me, "Copying RM Details ..."
  Dim wRsCustRM As MwfLib.MDORowSet
  '578.19 CrCtg= 'S' added
  Set wRsCustRM = moCn.OpenRes(" Select * from CustRM where (CrCtg= 'C' or CrCtg= 'S') and " + _
                 " CrCd= '" + ADC("wCustRmFr") + "' " + wCnd + " Order By CrRmCtg, CrOurRmCd, CrFrLn, CrToLn")
  With wRsCustRM
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying RM Details"
      GrdCustRm.AddItem
      wRow = GrdCustRm.Rows - 1
      GrdCustRm.Value(wRow, "CrRmCtg") = !CrRmCtg
      GrdCustRm.Value(wRow, "CrOurRmCd") = !CrOurRmCd
      GrdCustRm.Value(wRow, "CrFrLn") = !CrFrLn
      GrdCustRm.Value(wRow, "CrToLn") = !CrToLn
      GrdCustRm.Value(wRow, "CrCustRmCd") = !CrCustRmCd
      GrdCustRm.Value(wRow, "CrDesc") = !CrDesc
      GrdCustRm.Value(wRow, "CrPrtKey") = !CrPrtKey
      GrdCustRm.SaveRec (wRow)
      .MoveNext
    Loop
  End With
  MWLib.EndProcess Me
  GrdCustRm.EndCopy
  Set wRsCustRM = Nothing
  '*** End of the copy process
End Sub
Private Sub DispFra(ByVal pv_CustFra As en_CustFra)
  Select Case pv_CustFra
  Case Is = Cpy
    '*** If Frame FraCpy is visible then make it invisible and enable all the command buttons
    '*** If Frame FraCpy is invisible then make it visible and disable all the
        'command buttons except the one which is clicked. Change the background colour of
        'the clicked command button to white and make the font bold
    If FraCpy.Visible = True Then
      FraCpy.Visible = False
      FraCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpy)
      CmdCpy.SetFocus
    Else
      FraCpy.Visible = True
      FraCpy.Enabled = True
      ADC("wCustFr").SetFocus
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
    
  Case Is = CpyRm   'Sachin - 4.1.0.0
    If FraCpyRm.Visible = True Then
      FraCpyRm.Visible = False
      FraCpyRm.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRm)
      CmdCpyRm.SetFocus
    Else
      FraCpyRm.Visible = True
      FraCpyRm.Enabled = True
      ADC("wCustRmFr").SetFocus
      FraCpyRm.ZOrder
      Call EnaDisaCmds(True, CmdCpyRm)
    End If
    Case Is = ImpExc
    If FraImpExc.Visible = True Then
      FraImpExc.Visible = False
      FraImpExc.Enabled = False
      Call EnaDisaCmds(False, CmdImpExc)
      CmdImpExc.SetFocus
    Else
      FraImpExc.Visible = True
      FraImpExc.Enabled = True
      CmdImpExcPath.SetFocus
      FraImpExc.ZOrder
      Call EnaDisaCmds(True, CmdImpExc)
    End If
    '6.1-90 visible/invisible of scope frame
    Case Is = Scope
    If FraSC.Visible = True Then
      FraSC.Visible = False
      Call EnaDisaCmds(False, CmdSC)
      CmdSC.SetFocus
    Else
      FraSC.Visible = True
      ADC("wScDmCtgFr").SetFocus
      FraSC.ZOrder
      Call EnaDisaCmds(True, CmdSC)
    End If

  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** If second parameter is not given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false

  '*** If second parameter is not given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true

  '*** If second parameter is given and pv_ShowFra = True then the frame FraNKeyAll
      'on Frame Fra_nkeys is disabled and allowadd and allowdelete options of adc are
      'set to false and for the command button specified in the second parameter,
      'the BackColor is changed to white and the font is made Bold

  '*** If second parameter is given and pv_ShowFra = False then the frame FraNKeyAll
      'on Frame Fra_nkeys is enabled and allowadd and allowdelete options of adc are
      'set to true and for the command button specified in the second parameter,
      'the BackColor is changed to normal and the font is made Regular


  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = &H8000000F
        'pv_CmdBut.FontBold = False
      End If
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  ADC.AllowDelete = Not pv_ShowFra
  CmdCpy.Enabled = Not pv_ShowFra
  CmdCpyRm.Enabled = Not pv_ShowFra
  CmdImpExc.Enabled = Not pv_ShowFra
  CmdSC.Enabled = Not pv_ShowFra  '6.1-90
'Sachin 4.1.0.0
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
  
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpy.Visible = False
  FraCpy.Enabled = False
  FraCpyRm.Visible = False  'Sachin - 4.1.0.0
  FraImpExc.Visible = False 'Uni.9- Import from excel
  FraSC.Visible = False     '6.1-90

End Sub

Private Sub CmdImpExc_Click()
  Call DispFra(ImpExc)
  CmdSC.Enabled = False   '6.1-90 disable scope button
End Sub
'Uni.9- Getting excel file from selected path
Private Sub CmdImpExcPath_Click()
  CdImpExc.FileName = ""
  CdImpExc.DialogTitle = "Select Lab Rate Excel File "
  CdImpExc.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx"
  CdImpExc.FilterIndex = 1
  CdImpExc.ShowOpen
  ADC("WXLFILE") = CdImpExc.FileName

End Sub


'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  'FraNKeyAll.Height = FRA_NKEYS.Height + 500
  'FraNKeyAll.Width = FRA_NKEYS.Width + 500
  
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdCpyGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If ADC.HotKeys(KeyCode, Shift) Then
    KeyCode = 0
    Shift = 0
  End If
End Sub
Private Sub Form_Activate()
  Call FrmActivate(Me)
End Sub
Private Sub Form_Deactivate()
  Call FrmDeActivate(Me)
End Sub
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdCustDsg_Validate(Cancel As Boolean)
  Cancel = GrdCustDsg.Validate
End Sub
Private Sub GrdCustRm_Validate(Cancel As Boolean)
  Cancel = GrdCustRm.Validate
End Sub
Private Sub GrdCustAly_Validate(Cancel As Boolean)
  Cancel = GrdCustAly.Validate
End Sub
Private Sub GrdCustInst_Validate(Cancel As Boolean)
  Cancel = GrdCustInst.Validate
End Sub

Private Sub TXTADDR_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '578-20 once footer text boxes clicked other controls will disable
  ADC.AllowSave = False
  ADC.AllowDelete = False
  FraNKeyAll.Enabled = False
End Sub

'578.19-Store data from excel file to Grid
Private Sub ImportExcelRm()
    On Error GoTo ErrHdlr
    
    Dim wo_rsExcel As New MDORowSet, i As Integer
    Dim wCrCtg As String, wCrCd As String, wCrRmCtg As String, wCrRmCd As String
    Dim wRmFrLn As Double, wRmToLn As Double
    Dim wCrCustRmCd As String, wCrDesc As String

    'Checks proper excel file format
    If Not OpenExcelFile(ADC("WXLFILE")) Then DispMsg "Improper File Format", etError: Exit Sub
    
    'Checking proper sheet no.
    If ChkSheet(ADC("WXLSHTNO")) = False Then DispMsg "This WorkSheet Number Does Not Exist in the specified Excel File", etError: CloseExcel: Exit Sub
    
    'Creating Tmp file to store data
    moCn.CreateTmpTable "#TmpXls", "Select CrCtg ,CrCd, CrRmCtg, CrOurRmCd, CrFrLn, CrToLn, CrCustRmCd, CrDesc Into #TmpXls From CustRm Where 1=2"

    ADC("WXLFRROW") = IIF(ADC("WXLFRROW") <= 1, 2, ADC("WXLFRROW"))
    For i = ADC("WXLFRROW") To ADC("WXLTOROW")
        'Checking Customer code
        wCrCtg = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 1)))
        wCrCd = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 2)))
        
        If wCrCtg = "" Or wCrCd = "" Or wCrCtg <> ADC("CmCtg") Or wCrCd <> ADC("CmCd") Then _
          DispMsg "Customer code cannot be blank and needs to match the entry value. Cannot copy [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        'Checking valid RmCtg
        wCrRmCtg = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 3)))
        If Not moCn.RecSeek("Select 'x' From Param where PTyp= 'RMCTG' " + _
                             "and PMCd = '" + wCrRmCtg + "' ") Then _
                DispMsg "Cannot Import as Invalid Rm Ctg [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
    
        'Checking valid RmCode
        wCrRmCd = UCase(Trim(ReadCell(ADC("WXLSHTNO"), i, 4)))
        If Not moCn.RecSeek("Select 'x' from RmMst Where RmCd='" + wCrRmCd + "' And RmPrtKey='" + ctCurrPrtn + "' and RmValidYN='Y' and RmZ='N' ") Then _
          DispMsg "Cannot Import as Invalid RmCode [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
            
        'Checking Rm Ln/Sv/Ptr value
        If ReadCell(ADC("WXLSHTNO"), i, 5) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 5)) Then
            DispMsg "Cannot Copy as From Rm Ln/Sv/Ptr is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        If ReadCell(ADC("WXLSHTNO"), i, 6) <> "" And Not IsNumeric(ReadCell(ADC("WXLSHTNO"), i, 6)) Then
            DispMsg "Cannot Copy as To Rm Ln/Sv/Ptr is not Number [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        End If
        wRmFrLn = Val(ReadCell(ADC("WXLSHTNO"), i, 5))
        wRmToLn = Val(ReadCell(ADC("WXLSHTNO"), i, 6))
       
        If wRmFrLn < 0 Or wRmFrLn > 99999.9999 Then _
            DispMsg "Cannot Copy as Invalid From Rm Ln/Sv/Ptr, Should be Between 0 and 99999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wRmToLn < 0 Or wRmToLn > 99999.9999 Then _
            DispMsg "Cannot Copy as Invalid To Rm Ln/Sv/Ptr, Should be Between 0 and 99999.9999 [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If wRmFrLn > wRmToLn Then _
            DispMsg "Cannot Copy as From Rm Ln/Sv/Ptr <= To Rm Ln/Sv/Ptr [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
                
        wCrCustRmCd = UCase((ReadCell(ADC("WXLSHTNO"), i, 7)))
        If wCrCustRmCd = "" Then DispMsg "Cannot copy as Customer RmCode is Blank [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        If Len(wCrCustRmCd) > 20 Then DispMsg "Cannot copy as Customer Rm Code length is more than 20 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
        
        wCrDesc = UCase((ReadCell(ADC("WXLSHTNO"), i, 8)))
        If Len(wCrDesc) > 45 Then DispMsg "Cannot copy as Description is more than 45 characters [Row :" + CStr(i) + "]", etError: CloseExcel: Exit Sub
          
        moCn.Execute ("Insert Into #TmpXls Values (" + "'" + wCrCtg + "', '" + wCrCd + _
                      "', '" + wCrRmCtg + "','" + wCrRmCd + "'," + CStr(wRmFrLn) + "," + CStr(wRmToLn) + _
                        ",'" + wCrCustRmCd + "', '" + wCrDesc + "')")

    Next

    CloseExcel

    Set wo_rsExcel = moCn.OpenRes("Select * From #TmpXls ")
    Dim wRow As Double

    If wo_rsExcel.RecCount > 0 Then
      With GrdCustRm
        .StartCopy
        MWLib.BeginProcess Me, "Copying Details ..."
        Do While Not wo_rsExcel.EOF
          .AddItem
          wRow = .Rows - 1
          .Row = wRow
          .Value(wRow, "CrRmCtg") = wo_rsExcel!CrRmCtg
          .Value(wRow, "CrOurRmCd") = wo_rsExcel!CrOurRmCd
          .Value(wRow, "CrFrLn") = wo_rsExcel!CrFrLn
          .Value(wRow, "CrToLn") = wo_rsExcel!CrToLn
          .Value(wRow, "CrCustRmCd") = wo_rsExcel!CrCustRmCd
          .Value(wRow, "CrDesc") = wo_rsExcel!CrDesc

          If Not .SaveRec(wRow) Then
            DispMsg "Error trying to save " + wo_rsExcel!CrRmCtg + ":" + wo_rsExcel!CrOurRmCd, etError
          End If

          wo_rsExcel.MoveNext
        Loop

        MWLib.EndProcess Me
        .EndCopy
      End With
    End If

Exit Sub

ErrHdlr:
  If Err.Number = 429 Then
    DispMsg "MS Excel is either Corrupt Or Missing on this PC. Please Install/Re-Install MS Excel ! ", etError
  ElseIf Err.Number = 70 Then
    DispMsg "Cannot Access Template File. Close If Already Open and Continue.", etError
  Else
    DispMsg Err.Description, etError
  End If

End Sub
'6.1-90
Private Sub CmdSc_Click()
  Call DispFra(Scope)
End Sub


