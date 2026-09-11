VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed61.ocx"
Begin VB.Form EmrFrmDsgMst 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Master"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   14535
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdPDMLst 
      Height          =   465
      Left            =   3960
      TabIndex        =   231
      ToolTipText     =   "Option to Show PDM List"
      Top             =   7440
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "  Show &PDM List"
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
   Begin MwfCtl.MWCTL_BTN1 CmdMktPdf 
      Height          =   255
      Left            =   9495
      TabIndex        =   93
      ToolTipText     =   "Option To Open Market Card"
      Top             =   7665
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   450
      BackColor       =   -2147483633
      MaskColor       =   8454143
      DisableColor    =   0
      Caption         =   "Mkt Card"
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
      BtnLeftDiff     =   0
      BtnTopDiff      =   0
   End
   Begin MwfCtl.MWCTL_BTN1 CmdDsgMat 
      Height          =   465
      Left            =   10410
      TabIndex        =   94
      ToolTipText     =   "Option To Open Design Matrix"
      Top             =   7440
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Matrix View"
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
   Begin MwfCtl.MWCTL_BTN1 CmdPushCurr 
      Height          =   465
      Left            =   4860
      TabIndex        =   179
      ToolTipText     =   "Option To Copy Parts Raw Material"
      Top             =   7440
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Push &To Curr."
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyRm 
      Height          =   465
      Left            =   8475
      TabIndex        =   91
      ToolTipText     =   "Option To Copy Parts Raw Material"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy Part &RM "
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
   Begin MwfCtl.MWCTL_BTN1 CmdCpyFgBag 
      Height          =   465
      Left            =   7470
      TabIndex        =   90
      ToolTipText     =   "Option To Copy Details From Fg Bag"
      Top             =   7440
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Copy &Fg Bag"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSumm 
      Height          =   465
      Left            =   6585
      TabIndex        =   89
      ToolTipText     =   "Option To See The Summary Details"
      Top             =   7440
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Summary"
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
      Height          =   465
      Left            =   5790
      TabIndex        =   88
      ToolTipText     =   "Option To Copy Details From Other Design"
      Top             =   7440
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy"
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
      Left            =   10920
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7470
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   65
      Top             =   7455
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   979
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -30
      TabIndex        =   67
      Top             =   30
      Width           =   6525
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   4905
         TabIndex        =   2
         ToolTipText     =   "Enter Design Size/ Part"
         Top             =   15
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         DataField       =   "DmSz"
         IdName          =   "DMSZ"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         ToolTipText     =   "Enter Design Code"
         Top             =   0
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         DataField       =   "DmCd"
         IdName          =   "DMCD"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   165
         Index           =   58
         Left            =   5880
         TabIndex        =   180
         Top             =   0
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   291
         Enabled         =   0   'False
         DataType        =   4
         DataField       =   "DmPrtKey"
         IdName          =   "DMPRTKEY"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   165
         Index           =   2
         Left            =   6270
         TabIndex        =   181
         Top             =   15
         Visible         =   0   'False
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   291
         Enabled         =   0   'False
         MaxLength       =   3
         DataType        =   4
         DataField       =   "DmTcTyp"
         IdName          =   "DMTCTYP"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   66
         Left            =   1380
         TabIndex        =   0
         ToolTipText     =   "Enter Design Category"
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         ReCalcParent    =   "DMBAGPCS,DMUOM"
         IdName          =   "WDMCTG"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Left            =   2280
         TabIndex        =   213
         Top             =   0
         Width           =   135
      End
      Begin VB.Label LblDmSz 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sz"
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
         Left            =   3720
         TabIndex        =   70
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label LblDmCd 
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Ctg/Cd"
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
         Left            =   60
         TabIndex        =   68
         Top             =   30
         Width           =   1065
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7590
      Left            =   -120
      TabIndex        =   66
      Top             =   60
      Width           =   14625
      Begin MwfCtl.MWCTL_BTN1 CmdOpnPdf 
         Height          =   255
         Left            =   9615
         TabIndex        =   92
         ToolTipText     =   "Option To Open Job Card"
         Top             =   7365
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         BackColor       =   -2147483633
         MaskColor       =   8454143
         DisableColor    =   0
         Caption         =   "Job Card"
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
         BtnLeftDiff     =   0
         BtnTopDiff      =   0
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   7125
         Left            =   150
         TabIndex        =   69
         Top             =   30
         Width           =   14370
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   10380
            TabIndex        =   127
            Top             =   0
            Width           =   1815
            Begin VB.Image PicHsk 
               Height          =   1710
               Left            =   30
               Top             =   30
               Width           =   1710
            End
         End
         Begin VB.Frame FraPicDsgZoom 
            BorderStyle     =   0  'None
            Height          =   1755
            Left            =   9240
            TabIndex        =   220
            Top             =   0
            Width           =   1755
            Begin VB.Image PicDsgZoom 
               Height          =   1710
               Left            =   120
               Top             =   0
               Visible         =   0   'False
               Width           =   1710
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   9360
            TabIndex        =   219
            Top             =   0
            Width           =   1452
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   1305
            TabIndex        =   6
            ToolTipText     =   "Enter Design Lock (Y / N)"
            Top             =   1110
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "DmLockYN"
            IdName          =   "DmLockYN"
         End
         Begin VB.TextBox TxtPwd 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3510
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   1110
            Width           =   2340
         End
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   12360
            TabIndex        =   128
            Top             =   0
            Width           =   1815
            Begin VB.Image Pic3D 
               Height          =   1710
               Left            =   30
               Top             =   30
               Width           =   1710
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   30
            Left            =   1305
            TabIndex        =   3
            ToolTipText     =   "Enter Design Category"
            Top             =   270
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "DMCtg"
            ReCalcParent    =   "DMBAGPCS,DMUOM,DMPRCSSEQ"
            IdName          =   "DMCTG"
         End
         Begin TabDlg.SSTab TabDesgnDet 
            Height          =   4905
            Left            =   0
            TabIndex        =   10
            Top             =   2220
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   8652
            _Version        =   393216
            Tabs            =   12
            Tab             =   1
            TabsPerRow      =   12
            TabHeight       =   520
            TabCaption(0)   =   "<&1> General"
            TabPicture(0)   =   "EmrFrmDsgMst.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraDesgnDet(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "<&2>Bill Of Materials"
            TabPicture(1)   =   "EmrFrmDsgMst.frx":001C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "FraDesgnDet(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "<&3> Lab Details"
            TabPicture(2)   =   "EmrFrmDsgMst.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FraDesgnDet(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "<&4>Analysis"
            TabPicture(3)   =   "EmrFrmDsgMst.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "FraDesgnDet(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "<&5> History"
            TabPicture(4)   =   "EmrFrmDsgMst.frx":0070
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "FraDesgnDet(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "<&6> Faults"
            TabPicture(5)   =   "EmrFrmDsgMst.frx":008C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "FraDesgnDet(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "<&7> Comp Details"
            TabPicture(6)   =   "EmrFrmDsgMst.frx":00A8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "FraDesgnDet(6)"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "<&8>Model Avl."
            TabPicture(7)   =   "EmrFrmDsgMst.frx":00C4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "FraDesgnDet(7)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "<&9>Param"
            TabPicture(8)   =   "EmrFrmDsgMst.frx":00E0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "FraDesgnDet(8)"
            Tab(8).ControlCount=   1
            TabCaption(9)   =   "<&a> Images"
            TabPicture(9)   =   "EmrFrmDsgMst.frx":00FC
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "FraDesgnDet(9)"
            Tab(9).ControlCount=   1
            TabCaption(10)  =   "<&b> Prc Prd"
            TabPicture(10)  =   "EmrFrmDsgMst.frx":0118
            Tab(10).ControlEnabled=   0   'False
            Tab(10).Control(0)=   "FraDesgnDet(10)"
            Tab(10).ControlCount=   1
            TabCaption(11)  =   "<&d> Allow"
            TabPicture(11)  =   "EmrFrmDsgMst.frx":0134
            Tab(11).ControlEnabled=   0   'False
            Tab(11).Control(0)=   "FraDesgnDet(11)"
            Tab(11).ControlCount=   1
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4275
               Index           =   11
               Left            =   -74880
               TabIndex        =   229
               Top             =   480
               Width           =   14025
               Begin MwfCtl.MWCTL_FLX GrdDsgAllow 
                  Height          =   4230
                  Left            =   0
                  TabIndex        =   230
                  Top             =   0
                  Width           =   7155
                  _ExtentX        =   12621
                  _ExtentY        =   7461
                  Cols            =   6
                  colname1        =   "DMWTYP"
                  heading1        =   "Type          "
                  datafld1        =   "DmwTyp"
                  datatype1       =   4
                  recalcparent1   =   "GrdDsgAllow(DMWCMCD),GrdDsgAllow(DMWCMGRP),GrdDsgAllow(DMWREG),GrdDsgAllow(DMWYN)"
                  maxlength1      =   8
                  tooltiptext1    =   "Enter Type"
                  colname2        =   "DMWCMCD"
                  heading2        =   "Customer                      "
                  datafld2        =   "DmwCmCd"
                  datatype2       =   4
                  recalcon2       =   "GrdDsgAllow(DMWTYP)"
                  maxlength2      =   8
                  tooltiptext2    =   "Enter Customer Code"
                  colname3        =   "DMWCMGRP"
                  heading3        =   "Group                       "
                  datafld3        =   "DmwCmGrp"
                  datatype3       =   4
                  recalcon3       =   "GrdDsgAllow(DMWTYP)"
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Group"
                  colname4        =   "DMWREG"
                  heading4        =   "Region              "
                  datafld4        =   "DmwReg"
                  datatype4       =   4
                  recalcon4       =   "GrdDsgAllow(DMWTYP)"
                  maxlength4      =   5
                  tooltiptext4    =   "Enter Region"
                  colname5        =   "DMWYN"
                  heading5        =   "YN  "
                  datafld5        =   "DmwYN"
                  datatype5       =   4
                  recalcon5       =   "GrdDsgAllow(DMWTYP)"
                  maxlength5      =   1
                  tooltiptext5    =   "Enter Allow Yes/No"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4275
               Index           =   10
               Left            =   -74880
               TabIndex        =   209
               Top             =   480
               Width           =   14145
               Begin MwfCtl.MWCTL_FLX GrdPrcPrd 
                  Height          =   4095
                  Left            =   0
                  TabIndex        =   210
                  Top             =   0
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   7223
                  Cols            =   4
                  colname1        =   "DPPLOC"
                  heading1        =   "Prc                    "
                  datafld1        =   "DppLoc"
                  datatype1       =   4
                  maxlength1      =   8
                  tooltiptext1    =   "Process Code"
                  colname2        =   "DPPPRDCTG"
                  heading2        =   "Prd Ctg    "
                  datafld2        =   "DppPrdCtg"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Production Category Code"
                  colname3        =   "DPPPRTKEY"
                  heading3        =   "Prtn"
                  datafld3        =   "DppPrtKey"
                  maxlength3      =   1
                  style3          =   0
               End
            End
            Begin VB.Frame FraDesgnDet 
               Height          =   4185
               Index           =   9
               Left            =   -74820
               TabIndex        =   194
               Top             =   480
               Width           =   14025
               Begin VB.Frame FraPicDsg3 
                  BorderStyle     =   0  'None
                  Height          =   3240
                  Left            =   7920
                  TabIndex        =   218
                  Top             =   180
                  Width           =   3240
                  Begin VB.Image PicDsg3 
                     Height          =   2955
                     Left            =   120
                     Top             =   120
                     Width           =   2955
                  End
               End
               Begin VB.Frame FraPicDsg2 
                  BorderStyle     =   0  'None
                  Height          =   3240
                  Left            =   4080
                  TabIndex        =   217
                  Top             =   180
                  Width           =   3240
                  Begin VB.Image PicDsg2 
                     Height          =   2955
                     Left            =   120
                     Top             =   120
                     Width           =   2955
                  End
               End
               Begin VB.Frame FraPicDsg1 
                  BorderStyle     =   0  'None
                  Height          =   3240
                  Left            =   200
                  TabIndex        =   216
                  Top             =   180
                  Width           =   3240
                  Begin VB.Image PicDsg1 
                     Height          =   2955
                     Left            =   120
                     Top             =   120
                     Width           =   2955
                  End
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Image 1"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   7
                  Left            =   350
                  TabIndex        =   197
                  Top             =   3810
                  Width           =   885
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Image 2"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   66
                  Left            =   4220
                  TabIndex        =   196
                  Top             =   3780
                  Width           =   885
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Image 3"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   67
                  Left            =   8100
                  TabIndex        =   195
                  Top             =   3780
                  Width           =   885
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4395
               Index           =   8
               Left            =   -74970
               TabIndex        =   188
               Top             =   480
               Width           =   14235
               Begin VB.Frame FraPicDP 
                  BorderStyle     =   0  'None
                  Height          =   4335
                  Left            =   7440
                  TabIndex        =   189
                  Top             =   0
                  Width           =   4035
                  Begin VB.Image PicDP 
                     Height          =   4230
                     Left            =   30
                     Top             =   30
                     Width           =   3930
                  End
               End
               Begin MwfCtl.MWCTL_FLX GrdDsgPrm 
                  Height          =   4230
                  Left            =   60
                  TabIndex        =   190
                  Top             =   60
                  Width           =   3795
                  _ExtentX        =   6694
                  _ExtentY        =   7461
                  Cols            =   3
                  colname1        =   "DPTYP"
                  heading1        =   "Type     "
                  datafld1        =   "DpTyp"
                  datatype1       =   4
                  maxlength1      =   5
                  tooltiptext1    =   "Enter Design Param Type"
                  colname2        =   "DPCD"
                  heading2        =   "Code                                          "
                  datafld2        =   "DpCd"
                  datatype2       =   4
                  maxlength2      =   16
                  tooltiptext2    =   "Enter Code"
               End
               Begin MwfCtl.MWCTL_FLX GrdDsgPrmDet 
                  Height          =   4230
                  Left            =   4200
                  TabIndex        =   191
                  Top             =   60
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   7461
                  Cols            =   2
                  colname1        =   "DPDMCD"
                  heading1        =   "Design Code                     "
                  datafld1        =   "DpDmCd"
                  datatype1       =   4
                  maxlength1      =   15
                  tooltiptext1    =   "Enter Design Code"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4995
               Index           =   1
               Left            =   60
               TabIndex        =   149
               Top             =   480
               Width           =   14265
               Begin MwfCtl.MWCTL_FLX GrdDsgRm 
                  Height          =   4200
                  Left            =   0
                  TabIndex        =   46
                  Top             =   120
                  Width           =   14205
                  _ExtentX        =   25056
                  _ExtentY        =   7408
                  Cols            =   20
                  colname1        =   "DRSR"
                  heading1        =   "Sr.  "
                  datafld1        =   "DrSr"
                  datatype1       =   1
                  mask1           =   "##0"
                  maxlength1      =   3
                  style1          =   2
                  colname2        =   "DRRMCTG"
                  heading2        =   "Ctg"
                  datafld2        =   "DrRmCtg"
                  datatype2       =   4
                  recalcparent2   =   "GrdDsgRm(DRSUBSHP),GrdDsgRm(DRWSQTY)"
                  maxlength2      =   2
                  tooltiptext2    =   "Enter Raw Material Category"
                  colname3        =   "DRRMSCTG"
                  heading3        =   "SubCtg"
                  datafld3        =   "DrRmSCtg"
                  datatype3       =   4
                  recalcparent3   =   "GrdDsgRm(DRSUBSHP)"
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Raw Material Sub Category"
                  colname4        =   "DRRMCD"
                  heading4        =   "Rm Code                                  "
                  datafld4        =   "DrRmCd"
                  datatype4       =   4
                  recalcparent4   =   "GrdDsgRm(DRALYCD),GrdDsgRm(DRRMPTR),GrdDsgRm(DRPRDWT)"
                  maxlength4      =   16
                  tooltiptext4    =   "Enter Rm Code"
                  colname5        =   "DRLN1"
                  heading5        =   "Len/Sv"
                  datafld5        =   "DrLn1"
                  datatype5       =   2
                  mask5           =   "#0.000"
                  recalcparent5   =   "GrdDsgRm(DRRMPTR),GrdDsgRm(DRPRDWT)"
                  maxlength5      =   6
                  tooltiptext5    =   "Enter the Length or Sieve Sz of RmCode"
                  colname6        =   "DRLN2"
                  heading6        =   "Brdth  "
                  datafld6        =   "DrLn2"
                  datatype6       =   2
                  mask6           =   "#0.000"
                  recalcparent6   =   "GrdDsgRm(DRRMPTR),GrdDsgRm(DRPRDWT)"
                  maxlength6      =   6
                  tooltiptext6    =   "Enter the Breadth of RmCode"
                  colname7        =   "DRLN3"
                  heading7        =   "Dpth    "
                  datafld7        =   "DrLn3"
                  datatype7       =   2
                  mask7           =   "#0.000"
                  maxlength7      =   6
                  colname8        =   "DRPRDQTY"
                  heading8        =   "Prd Qty"
                  datafld8        =   "DrPrdQty"
                  datatype8       =   1
                  mask8           =   "###0"
                  recalcparent8   =   "GrdDsgRm(DRPRDWT),GrdDsgRm(DRQTY)"
                  maxlength8      =   4
                  tooltiptext8    =   "Enter Prd Qty for RmCode"
                  colname9        =   "DRRMPTR"
                  heading9        =   "Rm Ptr    "
                  datafld9        =   "DrRmPtr"
                  datatype9       =   2
                  mask9           =   "##0.0000"
                  recalcon9       =   "GrdDsgRm(DRRMCD),GrdDsgRm(DRLN1),GrdDsgRm(DRLN2)"
                  recalcparent9   =   "GrdDsgRm(DRPRDWT)"
                  maxlength9      =   8
                  tooltiptext9    =   "Enter Pointer Weight"
                  colname10       =   "DRPRDWT"
                  heading10       =   "Prd Wt     "
                  datafld10       =   "DrPrdWt"
                  datatype10      =   2
                  mask10          =   "-##0.000"
                  recalcon10      =   "GrdDsgRm(DRRMCD),GrdDsgRm(DRLN1),GrdDsgRm(DRLN2),GrdDsgRm(DRPRDQTY),GrdDsgRm(DRRMPTR)"
                  recalcparent10  =   "GrdDsgRm(DRWT)"
                  maxlength10     =   8
                  tooltiptext10   =   "Enter Prd Weight"
                  colname11       =   "DRQTY"
                  heading11       =   "Sal Qty"
                  datafld11       =   "DrQty"
                  datatype11      =   1
                  mask11          =   "###0"
                  recalcon11      =   "GrdDsgRm(DRPRDQTY)"
                  recalcparent11  =   "GrdDsgRm(DRWSQTY),GrdDsgRm(DRHSQTY)"
                  maxlength11     =   4
                  tooltiptext11   =   "Enter Qty for RmCode"
                  colname12       =   "DRWT"
                  heading12       =   "Sal Wt     "
                  datafld12       =   "DrWt"
                  datatype12      =   2
                  mask12          =   "-##0.000"
                  recalcon12      =   "GrdDsgRm(DRPRDWT)"
                  maxlength12     =   8
                  tooltiptext12   =   "Enter Weight for RmCode"
                  colname13       =   "DRSETSCD"
                  heading13       =   "Set Cd           "
                  datafld13       =   "DrSetSCd"
                  datatype13      =   4
                  maxlength13     =   8
                  tooltiptext13   =   "Enter Setting Sub Code"
                  colname14       =   "DRWSQTY"
                  heading14       =   "WSet"
                  datafld14       =   "DrWsQty"
                  datatype14      =   1
                  mask14          =   "###0"
                  recalcon14      =   "GrdDsgRm(DRQTY),GrdDsgRm(DRRMCTG)"
                  recalcparent14  =   "GrdDsgRm(DRHSQTY)"
                  maxlength14     =   4
                  tooltiptext14   =   "Enter Wax Set Quantity"
                  colname15       =   "DRHSQTY"
                  heading15       =   "HSet"
                  datafld15       =   "DrHsQty"
                  datatype15      =   1
                  mask15          =   "###0"
                  recalcon15      =   "GrdDsgRm(DRWSQTY),GrdDsgRm(DRQTY)"
                  maxlength15     =   4
                  style15         =   2
                  tooltiptext15   =   "Enter Hand Set Quantity"
                  colname16       =   "DRSUBSHP"
                  heading16       =   "SShp                "
                  datafld16       =   "DrSubShp"
                  datatype16      =   4
                  recalcon16      =   "GrdDsgRm(DRRMCTG),GrdDsgRm(DRRMSCTG)"
                  maxlength16     =   8
                  tooltiptext16   =   "Enter Sub Shape"
                  colname17       =   "DRALYCD"
                  heading17       =   "Alloy                                        "
                  datafld17       =   "DrAlyCd"
                  datatype17      =   4
                  recalcon17      =   "GrdDsgRm(DRRMCD)"
                  maxlength17     =   16
                  tooltiptext17   =   "Enter Alloy Code"
                  colname18       =   "DRMAINMET"
                  heading18       =   "MMet"
                  datafld18       =   "DrMainMet"
                  datatype18      =   4
                  maxlength18     =   1
                  tooltiptext18   =   "Enter Main Metal (Y/N)"
                  colname19       =   "DRPRTKEY"
                  heading19       =   "Prtn "
                  datafld19       =   "DrPrtKey"
                  maxlength19     =   1
                  style19         =   0
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4395
               Index           =   7
               Left            =   -74940
               TabIndex        =   167
               Top             =   480
               Width           =   14265
               Begin MwfCtl.MWCTL_FLX GrdDsgModAvl 
                  Height          =   4230
                  Left            =   60
                  TabIndex        =   64
                  Top             =   60
                  Width           =   8115
                  _ExtentX        =   14314
                  _ExtentY        =   7461
                  Cols            =   4
                  colname1        =   "DMASR"
                  heading1        =   "Sr.     "
                  datafld1        =   "DmaSr"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  colname2        =   "DMAAVLSZ"
                  heading2        =   "Avl. Sizes               "
                  datafld2        =   "DmaAvlSz"
                  datatype2       =   4
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Available Sizes"
                  colname3        =   "DMADESC"
                  heading3        =   "Description                                                                     "
                  datafld3        =   "DmaDesc"
                  maxlength3      =   45
                  tooltiptext3    =   "Enter Description"
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4275
               Index           =   6
               Left            =   -74940
               TabIndex        =   156
               Top             =   480
               Width           =   14265
               Begin MwfCtl.MWCTL_FLX GrdDsgCT 
                  Height          =   4215
                  Left            =   0
                  TabIndex        =   62
                  Top             =   0
                  Width           =   14145
                  _ExtentX        =   24950
                  _ExtentY        =   7435
                  Cols            =   14
                  colname1        =   "DCTSR"
                  heading1        =   "Sr     "
                  datafld1        =   "DctSr"
                  datatype1       =   1
                  mask1           =   "##0"
                  maxlength1      =   3
                  style1          =   2
                  tooltiptext1    =   "Serial Number"
                  colname2        =   "DCTCD"
                  heading2        =   "Comp Cd                     "
                  datafld2        =   "DctCd"
                  datatype2       =   4
                  maxlength2      =   15
                  tooltiptext2    =   "Enter Component Code"
                  colname3        =   "DCTSZ"
                  heading3        =   "Size    "
                  datafld3        =   "DctSz"
                  datatype3       =   4
                  maxlength3      =   5
                  tooltiptext3    =   "Enter Component Size"
                  colname4        =   "DCTCHR"
                  heading4        =   "Comp Chr"
                  datafld4        =   "DctChr"
                  datatype4       =   4
                  maxlength4      =   3
                  tooltiptext4    =   "Enter Component Character"
                  colname5        =   "DCTPARTS"
                  heading5        =   "Prt Qty"
                  datafld5        =   "DctParts"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Component Quantity"
                  colname6        =   "DCTRMCD"
                  heading6        =   "Main RmCd                             "
                  datafld6        =   "DctRmCd"
                  datatype6       =   4
                  recalcparent6   =   "GrdDsgCT(DCTRMWT)"
                  maxlength6      =   16
                  tooltiptext6    =   "Enter Main Raw Material Code"
                  colname7        =   "DCTDMCOL"
                  heading7        =   "Dm Color"
                  datafld7        =   "DctDmCol"
                  datatype7       =   4
                  maxlength7      =   5
                  tooltiptext7    =   "Enter Main Rm Color"
                  colname8        =   "DCTRMWT"
                  heading8        =   "Rm Weight"
                  datafld8        =   "DctRmWt"
                  datatype8       =   2
                  mask8           =   "###0.000"
                  recalcon8       =   "GrdDsgCT(DCTRMCD)"
                  maxlength8      =   8
                  tooltiptext8    =   "Enter Rm Weight"
                  colname9        =   "DCTDC"
                  heading9        =   "DC  "
                  datafld9        =   "DctDc"
                  datatype9       =   4
                  maxlength9      =   2
                  tooltiptext9    =   "Specify 'C' Only Clr Stn/ 'D' Only Diamond/ 'CD' for Both/ Blank if No Stone is used"
                  colname10       =   "DCTPRDSEQ"
                  heading10       =   "PrdSeq     "
                  datafld10       =   "DctPrdSeq"
                  datatype10      =   4
                  maxlength10     =   5
                  tooltiptext10   =   "Enter Production Sequence"
                  colname11       =   "DCTMAINPRTYN"
                  heading11       =   "MainPrt"
                  datafld11       =   "DctMainPrtYn"
                  datatype11      =   4
                  maxlength11     =   1
                  tooltiptext11   =   "Main Part Yes/ No"
                  colname12       =   "DCTDESC"
                  heading12       =   "Description                                                                                           "
                  datafld12       =   "DctDesc"
                  maxlength12     =   45
                  tooltiptext12   =   "Enter Discription for Compnent"
                  colname13       =   "DCTPRTKEY"
                  heading13       =   "Prtn "
                  datafld13       =   "DctPrtKey"
                  maxlength13     =   1
                  style13         =   0
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4275
               Index           =   5
               Left            =   -74880
               TabIndex        =   154
               Top             =   480
               Width           =   14025
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   3675
                  Index           =   46
                  Left            =   0
                  TabIndex        =   50
                  ToolTipText     =   "Enter Fault Description"
                  Top             =   480
                  Width           =   14085
                  _ExtentX        =   24844
                  _ExtentY        =   6482
                  Alignment       =   3
                  MaxLength       =   0
                  DataField       =   "DMFaults"
                  IdName          =   "DMFAULTS"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Faults"
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
                  Index           =   50
                  Left            =   90
                  TabIndex        =   155
                  Top             =   60
                  Width           =   1095
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4515
               Index           =   4
               Left            =   -74940
               TabIndex        =   152
               Top             =   360
               Width           =   14265
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   3915
                  Index           =   21
                  Left            =   120
                  TabIndex        =   49
                  ToolTipText     =   "Enter History Of The Design"
                  Top             =   360
                  Width           =   14085
                  _ExtentX        =   24844
                  _ExtentY        =   6906
                  Alignment       =   3
                  MaxLength       =   0
                  DataField       =   "DmHistory"
                  IdName          =   "DMHISTORY"
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
                  Index           =   21
                  Left            =   90
                  TabIndex        =   153
                  Top             =   60
                  Width           =   1095
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4395
               Index           =   3
               Left            =   -74940
               TabIndex        =   151
               Top             =   360
               Width           =   14265
               Begin MwfCtl.MWCTL_FLX GrdDsgAna 
                  Height          =   4230
                  Left            =   60
                  TabIndex        =   48
                  Top             =   60
                  Width           =   8835
                  _ExtentX        =   15584
                  _ExtentY        =   7461
                  Cols            =   6
                  colname1        =   "DAANASR"
                  heading1        =   "Sr. "
                  datafld1        =   "DaAnaSr"
                  datatype1       =   4
                  recalcparent1   =   "GrdDsgAna(WANANAME),GrdDsgAna(DAANACD),GrdDsgAna(DAANADESC)"
                  maxlength1      =   2
                  tooltiptext1    =   "Enter Analysis Serial No."
                  colname2        =   "WANANAME"
                  heading2        =   "Ana Name     "
                  recalcon2       =   "GrdDsgAna(DAANASR)"
                  maxlength2      =   10
                  style2          =   2
                  tooltiptext2    =   "Analysis Name"
                  colname3        =   "DAANACD"
                  heading3        =   "Ana Cd        "
                  datafld3        =   "DaAnaCd"
                  datatype3       =   4
                  recalcon3       =   "GrdDsgAna(DAANASR)"
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Analysis Code"
                  colname4        =   "DAANADESC"
                  heading4        =   "Analysis Desc                                                                        "
                  datafld4        =   "DaAnaDesc"
                  recalcon4       =   "GrdDsgAna(DAANASR)"
                  maxlength4      =   45
                  tooltiptext4    =   "Enter Analysis Description"
                  colname5        =   "DAPRTKEY"
                  heading5        =   "Prtn "
                  datafld5        =   "DaPrtKey"
                  maxlength5      =   1
                  style5          =   0
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4155
               Index           =   2
               Left            =   -74760
               TabIndex        =   150
               Top             =   600
               Width           =   13905
               Begin MwfCtl.MWCTL_FLX GrdDsgLab 
                  Height          =   3975
                  Left            =   0
                  TabIndex        =   47
                  Top             =   0
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   7011
                  Cols            =   7
                  colname1        =   "DLSR"
                  heading1        =   "Sr. "
                  datafld1        =   "DlSr"
                  datatype1       =   1
                  mask1           =   "#0"
                  maxlength1      =   2
                  style1          =   2
                  colname2        =   "DLMCD"
                  heading2        =   "Lab Cd "
                  datafld2        =   "DlMCd"
                  datatype2       =   4
                  recalcparent2   =   "GrdDsgLab(WQW)"
                  maxlength2      =   5
                  tooltiptext2    =   "Enter Labour Main Code"
                  colname3        =   "DLSCD"
                  heading3        =   "Lab SCd     "
                  datafld3        =   "DLSCd"
                  datatype3       =   4
                  recalcparent3   =   "GrdDsgLab(WQW)"
                  maxlength3      =   8
                  tooltiptext3    =   "Enter Labour Sub Code"
                  colname4        =   "WQW"
                  heading4        =   "By Q/W"
                  datatype4       =   4
                  recalcon4       =   "GrdDsgLab(DLMCD),GrdDsgLab(DLSCD)"
                  maxlength4      =   5
                  style4          =   2
                  colname5        =   "DLQTY"
                  heading5        =   "Qty"
                  datafld5        =   "DLQty"
                  datatype5       =   1
                  mask5           =   "##0"
                  maxlength5      =   3
                  tooltiptext5    =   "Enter Labour Quantity"
                  colname6        =   "DLPRTKEY"
                  heading6        =   "Prtn "
                  datafld6        =   "DlPrtKey"
                  maxlength6      =   1
                  style6          =   0
               End
            End
            Begin VB.Frame FraDesgnDet 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   4515
               Index           =   0
               Left            =   -74940
               TabIndex        =   11
               Top             =   360
               Width           =   14265
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   6
                  Left            =   13725
                  TabIndex        =   19
                  Top             =   285
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "##0"
                  MaxLength       =   3
                  DataType        =   1
                  DataField       =   "DMBagPcs"
                  ReCalcOn        =   "DMCTG"
                  IdName          =   "DMBAGPCS"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   8
                  Left            =   6135
                  TabIndex        =   21
                  ToolTipText     =   "Enter Old Design Code"
                  Top             =   610
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataType        =   4
                  DataField       =   "DMOldCd"
                  IdName          =   "DMOLDCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   9
                  Left            =   13335
                  TabIndex        =   22
                  ToolTipText     =   "Enter Colour"
                  Top             =   610
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMCol"
                  IdName          =   "DMCOL"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   10
                  Left            =   1455
                  TabIndex        =   27
                  ToolTipText     =   "Enter The Number Of Parts"
                  Top             =   930
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0"
                  MaxLength       =   4
                  DataType        =   1
                  DataField       =   "DMParts"
                  IdName          =   "DMPARTS"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   645
                  Index           =   11
                  Left            =   6135
                  TabIndex        =   28
                  ToolTipText     =   "Enter Description For Parts Of The Design"
                  Top             =   930
                  Width           =   8040
                  _ExtentX        =   14182
                  _ExtentY        =   1138
                  Alignment       =   3
                  MaxLength       =   90
                  DataField       =   "DMPartDesc"
                  IdName          =   "DMPARTDESC"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   12
                  Left            =   1455
                  TabIndex        =   29
                  ToolTipText     =   "Design On Hold Yes/No?"
                  Top             =   1215
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "DMHld"
                  ReCalcParent    =   "DMHLDDESC"
                  IdName          =   "DMHLD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   13
                  Left            =   6135
                  TabIndex        =   30
                  ToolTipText     =   "Enter The Reason For Keeping Design On Hold"
                  Top             =   1620
                  Width           =   8040
                  _ExtentX        =   14182
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "DMHldDesc"
                  ReCalcOn        =   "DMHLD"
                  IdName          =   "DMHLDDESC"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   870
                  Index           =   15
                  Left            =   6135
                  TabIndex        =   34
                  ToolTipText     =   "Enter Production Instruction"
                  Top             =   1920
                  Width           =   8040
                  _ExtentX        =   14182
                  _ExtentY        =   1535
                  Alignment       =   3
                  MaxLength       =   250
                  DataField       =   "DMPrdInst"
                  IdName          =   "DMPRDINST"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   16
                  Left            =   6135
                  TabIndex        =   37
                  ToolTipText     =   "Enter The Name/Code Of Designer"
                  Top             =   2790
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "DMDsgBy"
                  IdName          =   "DMDSGBY"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   19
                  Left            =   6135
                  TabIndex        =   39
                  ToolTipText     =   "Last Modified Date"
                  Top             =   3120
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   503
                  MaxLength       =   45
                  DataField       =   "DMLstMdf"
                  IdName          =   "DMLSTMDF"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   38
                  Left            =   6135
                  TabIndex        =   18
                  ToolTipText     =   "Enter Karatage"
                  Top             =   285
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMKt"
                  IdName          =   "DMKT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   42
                  Left            =   1455
                  TabIndex        =   20
                  ToolTipText     =   "Enter Set Family Code"
                  Top             =   610
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   503
                  MaxLength       =   15
                  DataType        =   4
                  DataField       =   "DmSetCd"
                  IdName          =   "DMSETCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   43
                  Left            =   13875
                  TabIndex        =   26
                  ToolTipText     =   "Specify Whether The Design Is Exclusive For The Customer For A Region Or For All Regions (Enter 'Y' or 'N')"
                  Top             =   4335
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "DmRegnExc"
                  ReCalcOn        =   "DMCMCD"
                  IdName          =   "DMREGNEXC"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   4
                  Left            =   1455
                  TabIndex        =   12
                  ToolTipText     =   "Enter Production Category"
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMPrdCtg"
                  IdName          =   "DMPRDCTG"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   5
                  Left            =   6135
                  TabIndex        =   13
                  ToolTipText     =   "Enter Sales Category"
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMSalCtg"
                  IdName          =   "DMSALCTG"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   41
                  Left            =   1455
                  TabIndex        =   17
                  ToolTipText     =   "Enter Loss Category"
                  Top             =   285
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DmLsCtg"
                  IdName          =   "DMLSCTG"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   39
                  Left            =   13320
                  TabIndex        =   16
                  ToolTipText     =   "Enter Value Addition Category"
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMVaCtg"
                  IdName          =   "DMVACTG"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   45
                  Left            =   9600
                  TabIndex        =   24
                  ToolTipText     =   "Enter Customer Name"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   165
                  _ExtentX        =   291
                  _ExtentY        =   503
                  MaxLength       =   45
                  ReCalcOn        =   "DMCMCD"
                  IdName          =   "WCMNAME"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   7
                  Left            =   1455
                  TabIndex        =   23
                  ToolTipText     =   "Enter Customer Code"
                  Top             =   4335
                  Visible         =   0   'False
                  Width           =   8295
                  _ExtentX        =   14631
                  _ExtentY        =   503
                  MaxLength       =   80
                  DataType        =   4
                  DataField       =   "DMCmCd"
                  ReCalcParent    =   "WCMREGNCD,DMREGNEXC,WCMNAME"
                  IdName          =   "DMCMCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   44
                  Left            =   9750
                  TabIndex        =   139
                  ToolTipText     =   "Customer Region Code"
                  Top             =   4335
                  Visible         =   0   'False
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   5
                  DataType        =   4
                  ReCalcOn        =   "DMCMCD"
                  IdName          =   "WCMREGNCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   14
                  Left            =   1455
                  TabIndex        =   31
                  ToolTipText     =   "Enter Production Sequence"
                  Top             =   1500
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DmDefSz"
                  IdName          =   "DMDEFSZ"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   17
                  Left            =   1455
                  TabIndex        =   36
                  ToolTipText     =   "Enter Design Date"
                  Top             =   2730
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "DMDsgDt"
                  IdName          =   "DMDSGDT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   18
                  Left            =   1455
                  TabIndex        =   38
                  ToolTipText     =   "Enter Model Maker"
                  Top             =   3015
                  Width           =   2355
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  MaxLength       =   20
                  DataField       =   "DMModMkr"
                  IdName          =   "DMMODMKR"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   51
                  Left            =   1455
                  TabIndex        =   40
                  ToolTipText     =   "Enter Wax Weight"
                  Top             =   3300
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  DataField       =   "DmWaxWt"
                  IdName          =   "DMWAXWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   52
                  Left            =   11805
                  TabIndex        =   41
                  ToolTipText     =   "Enter Model Weight With Runner"
                  Top             =   3480
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  DataField       =   "DmModRunWt"
                  IdName          =   "DMMODRUNWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   53
                  Left            =   11805
                  TabIndex        =   42
                  ToolTipText     =   "Enter Silver/Model Weight"
                  Top             =   3840
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  DataField       =   "DmSilModWt"
                  IdName          =   "DMSILMODWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   54
                  Left            =   1455
                  TabIndex        =   43
                  ToolTipText     =   "Enter Casting Piece Weight"
                  Top             =   3585
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "###0.000"
                  MaxLength       =   8
                  DataType        =   2
                  DataField       =   "DmCasPcWt"
                  IdName          =   "DMCASPCWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   57
                  Left            =   1455
                  TabIndex        =   44
                  ToolTipText     =   "Enter Total Diamond Weight (0 If Not To Be Checked)"
                  Top             =   3870
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "##0.000"
                  MaxLength       =   7
                  DataType        =   2
                  DataField       =   "DmTotDiaWt"
                  IdName          =   "DMTOTDIAWT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   59
                  Left            =   13305
                  TabIndex        =   182
                  ToolTipText     =   "Enter Design Date"
                  Top             =   2760
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
                  Enabled         =   0   'False
                  MaxLength       =   8
                  DataType        =   3
                  DataField       =   "DmCreatedDt"
                  IdName          =   "DMCREATEDDT"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   60
                  Left            =   1455
                  TabIndex        =   33
                  ToolTipText     =   "Enter Production Sequence"
                  Top             =   2130
                  Width           =   3000
                  _ExtentX        =   5292
                  _ExtentY        =   503
                  MaxLength       =   80
                  DataType        =   4
                  DataField       =   "DMPrdSeq"
                  IdName          =   "DMPRDSEQ"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   62
                  Left            =   6135
                  TabIndex        =   45
                  ToolTipText     =   "Enter Source Sketch/ Design Code"
                  Top             =   3630
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   503
                  MaxLength       =   15
                  DataType        =   4
                  DataField       =   "DmSrcDsgCd"
                  IdName          =   "DMSRCDSGCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   64
                  Left            =   1455
                  TabIndex        =   35
                  ToolTipText     =   "Enter Production Sequence"
                  Top             =   2430
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   503
                  MaxLength       =   8
                  DataType        =   4
                  DataField       =   "DmPrfVendCd"
                  IdName          =   "DMPRFVENDCD"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   65
                  Left            =   10920
                  TabIndex        =   25
                  ToolTipText     =   "Specify Whether The Design Exclusivity is  For The Customer Group (Y / N)"
                  Top             =   4320
                  Visible         =   0   'False
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   503
                  DataType        =   4
                  DataField       =   "DmCmGrp"
                  ReCalcOn        =   "DMCMCD"
                  IdName          =   "DMCMGRP"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   68
                  Left            =   1455
                  TabIndex        =   32
                  ToolTipText     =   "Enter Process Sequence"
                  Top             =   1800
                  Width           =   3000
                  _ExtentX        =   5292
                  _ExtentY        =   503
                  MaxLength       =   80
                  DataType        =   4
                  DataField       =   "DmPrcsSeq"
                  ReCalcOn        =   "DMCTG"
                  IdName          =   "DMPRCSSEQ"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   69
                  Left            =   8160
                  TabIndex        =   14
                  ToolTipText     =   "Enter Sales Category2"
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMSalCtg2"
                  IdName          =   "DMSALCTG2"
               End
               Begin MwfCtl.MWCTL_MED ATXT 
                  Height          =   285
                  Index           =   70
                  Left            =   10200
                  TabIndex        =   15
                  ToolTipText     =   "Enter Sales Category3"
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   503
                  MaxLength       =   5
                  DataType        =   4
                  DataField       =   "DMSalCtg3"
                  IdName          =   "DMSALCTG3"
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Ctg3"
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
                  Left            =   9120
                  TabIndex        =   228
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Ctg2"
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
                  Left            =   7080
                  TabIndex        =   227
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prc Seq"
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
                  Left            =   60
                  TabIndex        =   215
                  Top             =   1800
                  Width           =   1185
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Apply  to Cust Grp"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Index           =   69
                  Left            =   10680
                  TabIndex        =   212
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pref Vendor"
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
                  Left            =   60
                  TabIndex        =   211
                  Top             =   2400
                  Width           =   1485
               End
               Begin VB.Label LblSrcDsgCd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Source Sketch / Dsg Cd"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   570
                  Left            =   4650
                  TabIndex        =   192
                  Top             =   3630
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Seq"
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
                  Left            =   60
                  TabIndex        =   184
                  Top             =   2115
                  Width           =   1185
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
                  Index           =   62
                  Left            =   11820
                  TabIndex        =   183
                  Top             =   2760
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Dia Wt"
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
                  Left            =   60
                  TabIndex        =   178
                  Top             =   3870
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Casting Pc Wt"
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
                  Left            =   60
                  TabIndex        =   166
                  Top             =   3585
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Silver/Model Wt"
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
                  Left            =   9600
                  TabIndex        =   165
                  Top             =   3840
                  Width           =   1635
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Model Wt With Runner"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Index           =   56
                  Left            =   9600
                  TabIndex        =   164
                  Top             =   3480
                  Width           =   2415
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Wax Wt"
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
                  Left            =   60
                  TabIndex        =   163
                  Top             =   3300
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cust/ Region"
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
                  Left            =   60
                  TabIndex        =   138
                  Top             =   4335
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Ctg"
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
                  Left            =   4650
                  TabIndex        =   137
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Ctg"
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
                  TabIndex        =   136
                  Top             =   0
                  Width           =   855
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Loss Ctg"
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
                  Left            =   60
                  TabIndex        =   135
                  Top             =   285
                  Width           =   1005
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Val Addn Ctg"
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
                  Left            =   12000
                  TabIndex        =   134
                  Top             =   0
                  Width           =   1395
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Region Exc"
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
                  Left            =   12000
                  TabIndex        =   133
                  Top             =   4335
                  Visible         =   0   'False
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bag Pcs"
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
                  Left            =   12000
                  TabIndex        =   132
                  Top             =   285
                  Width           =   885
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Set Family Cd"
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
                  Left            =   60
                  TabIndex        =   131
                  Top             =   610
                  Width           =   1335
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Karat"
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
                  Left            =   4650
                  TabIndex        =   126
                  Top             =   285
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last Modified"
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
                  Left            =   4650
                  TabIndex        =   84
                  Top             =   3120
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Model Maker"
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
                  Left            =   60
                  TabIndex        =   83
                  Top             =   3015
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Design Date"
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
                  Left            =   60
                  TabIndex        =   82
                  Top             =   2730
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Designed By"
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
                  Left            =   4650
                  TabIndex        =   81
                  Top             =   2790
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prd Instruction"
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
                  Left            =   4650
                  TabIndex        =   80
                  Top             =   2280
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Default Size"
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
                  Left            =   60
                  TabIndex        =   79
                  Top             =   1470
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Holding Desc"
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
                  Left            =   4560
                  TabIndex        =   78
                  Top             =   1620
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "On Hold"
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
                  Left            =   60
                  TabIndex        =   77
                  Top             =   1185
                  Width           =   1365
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Part Desc"
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
                  Left            =   4650
                  TabIndex        =   76
                  Top             =   930
                  Width           =   1545
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Parts"
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
                  Left            =   60
                  TabIndex        =   75
                  Top             =   900
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Colour"
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
                  Left            =   12000
                  TabIndex        =   74
                  Top             =   610
                  Width           =   1485
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Old Cd"
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
                  Left            =   4650
                  TabIndex        =   73
                  Top             =   610
                  Width           =   1485
               End
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   570
            Index           =   3
            Left            =   1305
            TabIndex        =   5
            ToolTipText     =   "Enter Description"
            Top             =   555
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   1005
            Alignment       =   3
            MaxLength       =   45
            DataField       =   "DMDesc"
            IdName          =   "DMDESC"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   4845
            TabIndex        =   4
            ToolTipText     =   "Enter Unit Of Measurement"
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            DataField       =   "DMUOM"
            ReCalcOn        =   "DMCTG"
            IdName          =   "DMUOM"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   63
            Left            =   7245
            TabIndex        =   8
            ToolTipText     =   "Enter Design Valid (Y / N)"
            Top             =   1110
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            DataField       =   "DmValidYN"
            IdName          =   "DMVALIDYN"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   67
            Left            =   1290
            TabIndex        =   9
            ToolTipText     =   "Enter The Valid Co Cd's For Editing this Design"
            Top             =   1440
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   503
            MaxLength       =   45
            DataType        =   4
            DataField       =   "DmValidCoCd"
            IdName          =   "DMVALIDCOCD"
         End
         Begin VB.Label ALBL 
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
            Index           =   71
            Left            =   0
            TabIndex        =   214
            Top             =   1425
            Width           =   1245
         End
         Begin VB.Label LblValidYn 
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
            Left            =   6120
            TabIndex        =   208
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dsg Lock"
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
            Left            =   0
            TabIndex        =   148
            Top             =   1140
            Width           =   1005
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
            Height          =   285
            Index           =   20
            Left            =   2340
            TabIndex        =   147
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "3D Image"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   12900
            TabIndex        =   87
            Top             =   1950
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Sketch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   10830
            TabIndex        =   86
            Top             =   1950
            Width           =   885
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Uom"
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
            Left            =   4230
            TabIndex        =   85
            Top             =   270
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Desc"
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
            TabIndex        =   72
            Top             =   555
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ctg"
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
            Left            =   0
            TabIndex        =   71
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Frame FraSumm 
         Height          =   3885
         Left            =   1410
         TabIndex        =   103
         Top             =   3270
         Width           =   11775
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   26
            Left            =   1695
            TabIndex        =   104
            ToolTipText     =   "Gross Quantity"
            Top             =   1080
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WGRSQTY"
         End
         Begin MwfCtl.MWCTL_FLX GrdSumm 
            Height          =   3150
            Left            =   4410
            TabIndex        =   106
            Top             =   510
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   5556
            Cols            =   5
            colname1        =   "WRMCTG"
            heading1        =   "Ctg"
            datatype1       =   4
            maxlength1      =   2
            style1          =   2
            colname2        =   "WRMSCTG"
            heading2        =   "Sub Ctg"
            datatype2       =   4
            maxlength2      =   5
            style2          =   2
            colname3        =   "WQTY"
            heading3        =   "Qty     "
            datatype3       =   1
            mask3           =   "####0"
            maxlength3      =   5
            style3          =   2
            colname4        =   "WWT"
            heading4        =   "Wt.           "
            datatype4       =   2
            mask4           =   "-###0.000"
            maxlength4      =   9
            style4          =   2
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   27
            Left            =   2370
            TabIndex        =   108
            ToolTipText     =   "Gold Weight"
            Top             =   1365
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WGLDWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   28
            Left            =   2370
            TabIndex        =   110
            ToolTipText     =   "Other Metal Weight"
            Top             =   2505
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WOTHMTWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   29
            Left            =   2370
            TabIndex        =   112
            ToolTipText     =   "Diamond Weight"
            Top             =   2790
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WDIAWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   31
            Left            =   2370
            TabIndex        =   114
            ToolTipText     =   "Col Stone Weight"
            Top             =   3075
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WCSWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   32
            Left            =   2370
            TabIndex        =   116
            ToolTipText     =   "Accessory Weight"
            Top             =   3360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WACCWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   33
            Left            =   2370
            TabIndex        =   120
            ToolTipText     =   "Platinum Weight"
            Top             =   1650
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WPLWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   34
            Left            =   2370
            TabIndex        =   122
            ToolTipText     =   "Gross Weight"
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WGRSWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   35
            Left            =   1695
            TabIndex        =   123
            ToolTipText     =   "s"
            Top             =   2790
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WDIAQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   36
            Left            =   1695
            TabIndex        =   124
            ToolTipText     =   "Col Stone Quantity"
            Top             =   3075
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WCSQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   37
            Left            =   1695
            TabIndex        =   125
            ToolTipText     =   "Accessory Quantity"
            Top             =   3360
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "####0"
            MaxLength       =   5
            DataType        =   2
            IdName          =   "WACCQTY"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   40
            Left            =   1695
            TabIndex        =   129
            ToolTipText     =   "Karatage"
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WKT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   47
            Left            =   10605
            TabIndex        =   107
            ToolTipText     =   "Should The Production Rm Qty && Wt Be Shown In Dsg And Ord Entries"
            Top             =   3390
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WSHOWPRDRMDETS"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   50
            Left            =   2370
            TabIndex        =   159
            ToolTipText     =   "Platinum Weight"
            Top             =   1935
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WSILWT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   61
            Left            =   2370
            TabIndex        =   185
            ToolTipText     =   "Palladium Weight"
            Top             =   2220
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-###0.000"
            MaxLength       =   9
            DataType        =   2
            IdName          =   "WPDWT"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Palladium"
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
            Left            =   300
            TabIndex        =   187
            Top             =   2220
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   186
            Top             =   2235
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   162
            Top             =   1950
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   161
            Top             =   1650
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Silver"
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
            Left            =   300
            TabIndex        =   160
            Top             =   1935
            Width           =   1425
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Summary"
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
            TabIndex        =   157
            Top             =   90
            Width           =   11805
         End
         Begin VB.Label lblShBothWts 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Prd/ Sal wghts (Y/N)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7800
            TabIndex        =   146
            Top             =   3390
            Width           =   2745
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   145
            Top             =   2505
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "cts"
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
            Left            =   3390
            TabIndex        =   144
            Top             =   2790
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "cts"
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
            Index           =   46
            Left            =   3390
            TabIndex        =   143
            Top             =   3075
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Index           =   45
            Left            =   3390
            TabIndex        =   142
            Top             =   3360
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   141
            Top             =   1365
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "gms"
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
            Left            =   3390
            TabIndex        =   140
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Karat"
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
            Left            =   300
            TabIndex        =   130
            Top             =   510
            Width           =   855
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Platinum"
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
            Left            =   300
            TabIndex        =   121
            Top             =   1650
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Wt"
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
            Left            =   2670
            TabIndex        =   119
            Top             =   780
            Width           =   555
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
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
            Left            =   1830
            TabIndex        =   118
            Top             =   780
            Width           =   675
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Accessories"
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
            Left            =   300
            TabIndex        =   117
            Top             =   3360
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Col Stn"
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
            Left            =   300
            TabIndex        =   115
            Top             =   3075
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Dia"
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
            Left            =   300
            TabIndex        =   113
            Top             =   2790
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Oth Mt"
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
            Index           =   28
            Left            =   300
            TabIndex        =   111
            Top             =   2505
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gold"
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
            Left            =   300
            TabIndex        =   109
            Top             =   1365
            Width           =   1425
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross"
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
            Left            =   300
            TabIndex        =   105
            Top             =   1080
            Width           =   1425
         End
      End
      Begin VB.Frame FraPrtRmCpy 
         Height          =   4125
         Left            =   7320
         TabIndex        =   168
         Top             =   2640
         Width           =   4125
         Begin MwfCtl.MWCTL_BTN1 CmdPrtRmCpyGo 
            Height          =   390
            Left            =   3480
            TabIndex        =   169
            ToolTipText     =   "Click Here To Copy"
            Top             =   3600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   688
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
         Begin MwfCtl.MWCTL_FLX GrdDsgTgt 
            Height          =   3540
            Left            =   90
            TabIndex        =   170
            Top             =   450
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   6244
            Cols            =   5
            colname1        =   "DTSR"
            heading1        =   "Sr"
            datafld1        =   "DtSr"
            datatype1       =   1
            mask1           =   "#0"
            maxlength1      =   2
            style1          =   0
            colname2        =   "DTRMCTG"
            heading2        =   "Ctg"
            datafld2        =   "DtRmCtg"
            datatype2       =   4
            maxlength2      =   2
            tooltiptext2    =   "Enter Raw Material Category"
            colname3        =   "DTRMSCTG"
            heading3        =   "SubCtg"
            datafld3        =   "DtRmSCtg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Enter Raw Material Sub Category"
            colname4        =   "DTRMCD"
            heading4        =   "Rm Code             "
            datafld4        =   "DtRmCd"
            datatype4       =   4
            maxlength4      =   12
            tooltiptext4    =   "Enter Target Raw Material Code"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Change Rm With"
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
            Index           =   2
            Left            =   -30
            TabIndex        =   171
            Top             =   90
            Width           =   4125
         End
      End
      Begin VB.Frame FraCpy 
         Height          =   1665
         Left            =   5715
         TabIndex        =   95
         Top             =   5490
         Width           =   3765
         Begin MwfCtl.MWCTL_BTN1 CmdCpyGo 
            Height          =   495
            Left            =   2985
            TabIndex        =   102
            ToolTipText     =   "Click Here To Copy"
            Top             =   1050
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   23
            Left            =   1185
            TabIndex        =   96
            ToolTipText     =   "Enter Source Design Type"
            Top             =   420
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WDMTCTYP"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   24
            Left            =   1185
            TabIndex        =   98
            ToolTipText     =   "Enter Source Design Code"
            Top             =   705
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WDMCD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   25
            Left            =   1185
            TabIndex        =   100
            ToolTipText     =   "Enter Source Design Size/Part"
            Top             =   990
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDMSZ"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From"
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
            TabIndex        =   158
            Top             =   90
            Width           =   3765
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Size / Part"
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
            Left            =   60
            TabIndex        =   101
            Top             =   990
            Width           =   1275
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Desgn Cd"
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
            Left            =   60
            TabIndex        =   99
            Top             =   705
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "TcType"
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
            Left            =   60
            TabIndex        =   97
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame FraPDMList 
         Height          =   5295
         Left            =   1920
         TabIndex        =   232
         Top             =   2040
         Width           =   10605
         Begin MwfCtl.MWCTL_FLX GrdPDMLst 
            Height          =   5115
            Left            =   30
            TabIndex        =   233
            Top             =   120
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   9022
            Cols            =   7
            colname1        =   "WPDMVCHR"
            heading1        =   "PDM Voucher                           "
            datatype1       =   4
            maxlength1      =   30
            style1          =   2
            colname2        =   "WPDMCMCTG"
            heading2        =   "CmCtg"
            datatype2       =   4
            maxlength2      =   2
            style2          =   2
            colname3        =   "WPDMCMCD"
            heading3        =   "Cust Code          "
            datatype3       =   4
            maxlength3      =   8
            style3          =   2
            colname4        =   "WPDMDMCD"
            heading4        =   "Design Code                            "
            datatype4       =   4
            maxlength4      =   15
            style4          =   2
            colname5        =   "WPDMSFX"
            heading5        =   "Suffix                         "
            datatype5       =   4
            maxlength5      =   10
            style5          =   2
            colname6        =   "WPDMSZ"
            heading6        =   "Size         "
            datatype6       =   4
            maxlength6      =   5
            style6          =   2
         End
      End
      Begin VB.Frame FraCpyFgBag 
         Height          =   5655
         Left            =   4470
         TabIndex        =   172
         Top             =   1500
         Width           =   6885
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   226
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton OptFgbag 
               Caption         =   "FG Bag"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   51
               ToolTipText     =   "Copy from FG Bag"
               Top             =   160
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptMstRec 
               Caption         =   "Master Record"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   52
               ToolTipText     =   "Copy from Master Record"
               Top             =   480
               Width           =   2055
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "ForceCopy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   225
            Top             =   4800
            Width           =   4455
            Begin VB.CheckBox ChkFCSTW 
               Caption         =   "Pointer Weights from STW Chart"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   60
               ToolTipText     =   "Check for Pointer Weights from STW Chart"
               Top             =   240
               Width           =   3135
            End
         End
         Begin VB.Frame FraDsgGrdFil 
            Height          =   975
            Left            =   120
            TabIndex        =   222
            Top             =   3840
            Width           =   4455
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   333
               Left            =   1560
               TabIndex        =   57
               ToolTipText     =   "Enter Group Option For Rm Summary"
               Top             =   120
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WCPYFGBAGGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   334
               Left            =   1560
               TabIndex        =   58
               ToolTipText     =   "Enter Raw Material Category "
               Top             =   480
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               IdName          =   "WCPYFGBAGCTG"
            End
            Begin MwfCtl.MWCTL_BTN1 CmdCpyFgBagFill 
               Height          =   375
               Left            =   3480
               TabIndex        =   59
               ToolTipText     =   "Click Here To Fill Grid"
               Top             =   360
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
               BackColor       =   -2147483633
               MaskColor       =   8454143
               Caption         =   "Fill G&rid"
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
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Group By"
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
               Index           =   259
               Left            =   120
               TabIndex        =   224
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Cateogry"
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
               Index           =   260
               Left            =   120
               TabIndex        =   223
               Top             =   480
               Width           =   1365
            End
         End
         Begin MwfCtl.MWCTL_BTN1 CmdCpyFgBagGo 
            Height          =   495
            Left            =   5520
            TabIndex        =   61
            ToolTipText     =   "Click Here To Copy"
            Top             =   4440
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   3870
            TabIndex        =   53
            ToolTipText     =   "Enter Source Company Code"
            Top             =   495
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WCOCD"
            CmpStr          =   "BCoCd="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   49
            Left            =   3870
            TabIndex        =   54
            ToolTipText     =   "Enter Fg Bag Yy"
            Top             =   780
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WBYY"
            CmpStr          =   "BYy="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   55
            Left            =   4350
            TabIndex        =   55
            ToolTipText     =   "Enter Fg Bag Character"
            Top             =   780
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            MaxLength       =   3
            DataType        =   4
            IdName          =   "WBCHR"
            CmpStr          =   "BChr="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   56
            Left            =   5145
            TabIndex        =   56
            ToolTipText     =   "Enter Fg Bag Number"
            Top             =   780
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WBNO"
            CmpStr          =   "BNo="
         End
         Begin MwfCtl.MWCTL_FLX GrdCpyFgBag 
            Height          =   2475
            Left            =   120
            TabIndex        =   221
            Top             =   1320
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   4366
            Cols            =   7
            colname1        =   "DTSR"
            heading1        =   "Sr."
            datafld1        =   "DtSr"
            datatype1       =   1
            mask1           =   "###0"
            maxlength1      =   4
            style1          =   0
            colname2        =   "DTRMCTG"
            heading2        =   "Src Ctg"
            datafld2        =   "DtRmCtg"
            datatype2       =   4
            maxlength2      =   2
            tooltiptext2    =   "Enter Source Raw Material Category"
            colname3        =   "DTRMSCTG"
            heading3        =   "Src SubCtg"
            datafld3        =   "DtRmSCtg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Enter Source Raw Material SubCategory"
            colname4        =   "DTCOL"
            heading4        =   "Src Color"
            datafld4        =   "DtCol"
            datatype4       =   4
            maxlength4      =   5
            tooltiptext4    =   "Enter Source Raw Material Colour"
            colname5        =   "DTRMRSZ"
            heading5        =   "Src RangeSz"
            datafld5        =   "DtRmRSZ"
            datatype5       =   4
            maxlength5      =   12
            tooltiptext5    =   "Enter Source Raw Material RangeSize"
            colname6        =   "DTRMCD"
            heading6        =   "Tgt Rm Code                            "
            datafld6        =   "DtRmcd"
            datatype6       =   4
            maxlength6      =   16
            tooltiptext6    =   "Enter Target Raw Material Code"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   5040
            TabIndex        =   177
            Top             =   780
            Width           =   75
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Comp Cd"
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
            Left            =   2760
            TabIndex        =   176
            Top             =   495
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "BYy/Chr/No"
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
            Left            =   2760
            TabIndex        =   175
            Top             =   780
            Width           =   1125
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   4260
            TabIndex        =   174
            Top             =   780
            Width           =   75
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From Finished Goods Bag"
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
            TabIndex        =   173
            Top             =   90
            Width           =   6795
         End
      End
      Begin VB.OLE OLE1 
         Class           =   "AcroExch.Document.7"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   193
         Top             =   7200
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame FraDsgMatrix 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   0
      TabIndex        =   198
      Top             =   0
      Width           =   8055
      Begin VB.Image ImgMatZoom 
         Height          =   7605
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   7785
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   205
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   204
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   5400
         TabIndex        =   203
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   5400
         TabIndex        =   206
         Top             =   4960
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   2760
         TabIndex        =   207
         Top             =   4960
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   202
         Top             =   4960
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   5400
         TabIndex        =   200
         Top             =   7530
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   201
         Top             =   7530
         Width           =   2535
      End
      Begin VB.Label LblPic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2760
         TabIndex        =   199
         Top             =   7530
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   8
         Left            =   2760
         Top             =   5370
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   9
         Left            =   5400
         Top             =   5370
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   6
         Left            =   5400
         Top             =   2800
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   5
         Left            =   2760
         Top             =   2800
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   7
         Left            =   120
         Top             =   5370
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   4
         Left            =   120
         Top             =   2800
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   1
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00FFFFFF&
         Height          =   2460
         Index           =   2
         Left            =   2760
         Top             =   240
         Width           =   2535
      End
      Begin VB.Shape Shape 
         Height          =   2460
         Index           =   3
         Left            =   5400
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   6
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   4
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2200
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2200
         Index           =   2
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2200
         Index           =   3
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   5
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2800
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   7
         Left            =   120
         Stretch         =   -1  'True
         Top             =   5370
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   8
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   5370
         Width           =   1650
      End
      Begin VB.Image ImgArr 
         Appearance      =   0  'Flat
         Height          =   2205
         Index           =   9
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   5370
         Width           =   1650
      End
   End
End
Attribute VB_Name = "EmrFrmDsgMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Entry --------'
' 3 Menus 'DM', 'PM', 'SM' all similar, except
'   in 'PM', DmSz <> '', and 'DM' record to exist for DmSz='' (checked in trigger)
'
' X-DsgMst, A-DsgRm, D-DsgLab.
'
' TabPage, TabDesgnDet, 5 Tabs
'   1. General
'   2. BOM, GrdDsgRm
'   3. Lab Details, GrdDsgLab
'   4. Design Analysis Details (Currently allowed only for Dsg TcTyp= 'DM' and DmSz= '')
'   5. History
'   6. Faults
' CmdCpy to copy from another design, wDmTcTyp, wDmCd, wDmSz, FraCpy, CmdCpyGo
'     private sub CpyDesgDets
' CmdSumm, FraSumm, GrdSumm is filled, CmdSummGo
'            wGrsQty, wGrsWt, wGldWt,.....wAccWt are displayed
'             private sub Summary
'
' Picture Button  PicSk for the pencil Sketch
'                 Pic3d for the 3-D image.
' Enlarged state of the pictures can be obtained by double clicking on the pictures
' The pictures can be brought back to the original size by again double clicking the pictures
'
' Negative values are allowed in the Weight field
'--------------------
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mb_Upd As Boolean, ws_MultiCoMod As Boolean
'*** Declare variables wPic3DHt, wPic3DWdt, wPicHskHt, wPicHskWdt as single to store the
    'original Height and Width of the 3D and Sketch Picture controls
'*** Declare variables wFraPic3DHt, wFraPic3DWdt, wFraPicHskHt, wFraPicHskWdt,
    'wFraPic3DLft, wFraPic3DTop, wFraPicHskLft, wFraPicHskTop as single
    'to store the original Sizes & Positions of the 3D and Sketch Picture controls
'*** Declare variables wHskPth and w3DPth to store the picture paths for Sketch And 3D Images

Dim wPic3DHt As Single
Dim wPic3DWdt As Single
Dim wPicHskHt As Single
Dim wPicHskWdt As Single
Dim ms_hAddJROYN As String
Dim wPicDPHt As Single, wPicDPWdt As Single     '*** Jay 3.4[DP]

Dim wFraPic3DHt As Single
Dim wFraPic3DWdt As Single
Dim wFraPicHskHt As Single
Dim wFraPicHskWdt As Single

Dim wFraPic3DLft As Single
Dim wFraPic3DTop As Single
Dim wFraPicHskLft As Single
Dim wFraPicHskTop As Single
Dim wPic3DDblClk As Boolean
Dim wPicHskDblClk As Boolean

Dim wPicDsgZoomHt As Single
Dim wPicDsgZoomWdt As Single
Dim wFraPicDsgZoomHt As Single
Dim wFraPicDsgZoomWdt As Single
Dim wFraPicDsgZoomLft As Single
Dim wFraPicDsgZoomTop As Single

'*** Jay 2.14 *** (SEO Change)(wPwdFlg Changed)
Dim ms_PwdFlag As Boolean
'Dim ms_Pwd As String
'*** Jay 2.14 *** (SEO Change)(wPwdFlg Changed)

Dim wHskPth As String
Dim w3DPth As String

Dim wDsg1Pth As String
Dim wDsg2Pth As String
Dim wDsg3Pth As String

Dim wShBothWts As String, wModAvl As String

Dim ms_LockDt As String
Dim mf_DtLocked As Boolean
Dim ms_FutureDt As String

'*** For the Toggle Options of Command Buttons
Enum en_DsgFra
  Cpy = 0
  Summ = 1
  Pwd = 2
  CpyRm = 3
  CpyFgBag = 4
  DsgMat = 5        '4.1.2
  PDMLst = 6    '6.1b-46
End Enum

'*** Jay 2.13(CT) ***
Dim wb_CTMod As Boolean
'*** Jay 2.13(CT) ***

'****** Sachin 3.3.0 (25-06-08) - [28. Valid RM Size Master ] ******
Dim mb_VldRmLn As Boolean
'****** Sachin 3.3.0 (25-06-08) - [28. Valid RM Size Master ] ******

Dim mb_ActCalcWt As Boolean     '3.11.0

' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
Dim mb_DsgCat As Boolean
' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele

Dim mOldWidth As Single, mOldHt As Single, mOldLeft As Single, mOldTop As Single
Private Sub ADC_Load()
  '*** Set the First Key Control property to Design Code
  '*** Set the First Non Key Control property to Design Category
  '*** Set the Child Property Of ADC to GrdDsgRm, GrdDsgLab
  '*** Set the Previous and Next Control Properties for the Grids GrdDsgRm and GrdDsgLab
  '*** Set the Hot Keys for the Grids GrdDsgRm and GrdDsgLab as 'B' and 'C'
  '*** Set the current tab as 0
  '*** Set the caption of the form according to the menu option selected by the user

  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  Set ADC.FirKeyCtl = ADC("wDmCtg")
 'Set ADC.FirNKeyCtl = CmdPwd
  Set ADC.FirNKeyCtl = ADC("DmCtg")
  '*** Jay 2.13(CT) *** (GrdDsgCT ADDED)
  '*** Jay 3.4[DP]
  'ADC.Child = "GrdDsgRm,GrdDsgLab,GrdDsgAna,GrdDsgCT,GrdDsgTgt,GrdDsgModAvl,GrdDsgPrm"
  ADC.Child = "GrdDsgRm,GrdDsgLab,GrdDsgAna,GrdDsgCT,GrdDsgTgt,GrdDsgModAvl,GrdDsgPrm,GrdPrcPrd,GrdCpyFgBag,GrdDsgAllow"   '****Bhavna added GrdPRcPrd
  GrdDsgPrm.Child = "GrdDsgPrmDet"
  '*** Jay 3.4[DP]
    
  Set GrdDsgRm.PrevCtl = ADC("DmDesc")
  Set GrdDsgRm.NextCtl = ADC
  Set GrdDsgLab.PrevCtl = ADC("DmDesc")
  Set GrdDsgLab.NextCtl = ADC
  Set GrdSumm.PrevCtl = CmdSumm
  Set GrdSumm.NextCtl = ADC("wShowPrdRmDets")
  
  '*** Geeta***2.10
  Set GrdDsgTgt.PrevCtl = CmdCpyRm
  Set GrdDsgTgt.NextCtl = CmdPrtRmCpyGo
  '*** Jay 2.13(CT) ***
  Set GrdDsgCT.PrevCtl = ADC("DmDesc")
  Set GrdDsgCT.NextCtl = ADC
  '*** Jay 2.13(CT) ***
  '*** Jay 3.4[DP]
  Set GrdDsgPrm.PrevCtl = ADC
  Set GrdDsgPrm.NextCtl = ADC
  Set GrdDsgPrmDet.PrevCtl = GrdDsgPrm
  Set GrdDsgPrmDet.NextCtl = GrdDsgPrm
  '*** Jay 3.4[DP]
  '*****bhavna
  Set GrdPrcPrd.PrevCtl = ADC("DmDesc")
  Set GrdPrcPrd.NextCtl = ADC
  '*****bhavna
    
  GrdDsgRm.HotKey = "B"
  GrdDsgLab.HotKey = "C"
  GrdDsgAna.HotKey = "D"
  '*** Jay 3.4[DP]
  GrdDsgPrm.HotKey = "E"
  GrdDsgPrmDet.HotKey = "F"
  '*** Jay 3.4[DP]
  
  'FraNKeyAll.ZOrder (0)
  TabDesgnDet.Tab = 0
  If ADC.MenuCd = "DM" Then Me.Caption = "Design Master"
  If ADC.MenuCd = "SM" Then Me.Caption = "Design Sketch Master"
  
  '*** Jay 2.13(CT) ***
  If ADC.MenuCd = "PM" Then Me.Caption = "Design Component Master": LblDmCd.Caption = "Comp Cd"
  'If ADC.MenuCd = "PM" Then Me.Caption = "Design Part Master": LblDmCd.Caption = "Part Cd": LblDmSz.Caption = "Sz"
  '#### IF wb_CTMod = True  then module = Y else module = N
  wb_CTMod = moCn.RecSeek("Select hCompBagsYn From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' " + _
                            " and hCompBagsYn = 'Y'")
  '*** Jay 2.13(CT) ***
  
  '****** Sachin 2.13.0 - 25/06/2006 ******
  wModAvl = moCn.GetFldVal("Select HModAvlYN From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'")
  '****** Sachin 2.13.0 - 25/06/2006 ******
  
  '3.11.0
  mb_ActCalcWt = moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' And hCpyActCalcWtYN='Y'")
    
  '****** Sachin 3.02 [26/11/07] - Partition Key Condition ******
  CmdPushCurr.Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  '****** Sachin 3.02 [26/11/07] - Partition Key Condition ******
  
  ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
  mb_DsgCat = moCn.RecSeek("Select 'x' from Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and HDsgCatYn='Y' ")
  ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
   
  If ADC.MenuCd = "DM" Then
    'TabDesgnDet.TabsPerRow = 8
    '****** Sachin 2.13.0 - 25/06/2006 ******
    If wModAvl = "Y" Then
      TabDesgnDet.TabVisible(7) = True
      TabDesgnDet.TabEnabled(7) = True
    ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Module
      If mb_DsgCat = True Then
        TabDesgnDet.TabVisible(8) = True: TabDesgnDet.TabEnabled(8) = True
        TabDesgnDet.TabsPerRow = 12   '578-12 changed 11 to 12   '*** Jay 3.4[DP]
      Else
        TabDesgnDet.TabVisible(8) = False: TabDesgnDet.TabEnabled(8) = False
        TabDesgnDet.TabsPerRow = 11   '578-12 changed 10 to 11    '*** Jay 3.4[DP]
      End If
    ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Module
   Else
      TabDesgnDet.TabVisible(7) = False
      TabDesgnDet.TabEnabled(7) = False
      ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Module
      If mb_DsgCat = True Then
        TabDesgnDet.TabVisible(8) = True: TabDesgnDet.TabEnabled(8) = True
        TabDesgnDet.TabsPerRow = 12   '578-12 changed 11 to 12  '*** Jay 3.4[DP]
      Else
        TabDesgnDet.TabVisible(8) = False: TabDesgnDet.TabEnabled(8) = False
        TabDesgnDet.TabsPerRow = 11   '578-12 changed 10 to 11  '*** Jay 3.4[DP]
      End If
      ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Module
   End If
'    '****** Sachin 2.13.0 - 25/06/2006 ******
    CmdDsgMat.Visible = True        '4.1.2


  ElseIf ADC.MenuCd = "SM" Or ADC.MenuCd = "PM" Then
    TabDesgnDet.TabVisible(3) = True
    TabDesgnDet.TabEnabled(3) = True
    '****** Sachin 2.13.0 - 02-06-2006 ******
    TabDesgnDet.TabVisible(7) = False
    TabDesgnDet.TabEnabled(7) = False
    '****** Sachin 2.13.0 - 02-06-2006 ******
    '*** Jay 2.13(CT) ***
    TabDesgnDet.TabVisible(6) = False
    TabDesgnDet.TabEnabled(6) = False
    '*** Jay 3.4[DP]
    TabDesgnDet.TabVisible(8) = False
    TabDesgnDet.TabEnabled(8) = False
    '*** Jay 3.4[DP]
    TabDesgnDet.TabsPerRow = 9    '578-12 changed 8 to 9
    'TabDesgnDet.TabsPerRow = 6
    '*** Jay 2.13(CT) ***
    
    '4.1.3.0
    ADC("DmValidYn") = "Y"
    LblValidYn.Visible = False
    ADC("DmValidYn").Visible = False
  End If
  '********bhavna 5.1.0
    If ADC.MenuCd = "DM" Or ADC.MenuCd = "SM" Then
        TabDesgnDet.TabVisible(10) = True
        TabDesgnDet.TabEnabled(10) = True
        TabDesgnDet.TabVisible(11) = True   '578-12 allow tab visible only for DM,SM type
        TabDesgnDet.TabEnabled(11) = True   '578-12 allow tab visible only for DM,SM type
    Else
        TabDesgnDet.TabVisible(10) = False
        TabDesgnDet.TabEnabled(10) = False
        TabDesgnDet.TabVisible(11) = False   '578-12 allow tab visible only for DM,SM type
        TabDesgnDet.TabEnabled(11) = False   '578-12 allow tab visible only for DM,SM type
    End If
  '********bhavna 5.1.0
  
  '*** Storing The Original Size And Position Of The Picture Controls ***
  wPic3DHt = Pic3D.Height
  wPic3DWdt = Pic3D.Width
  wPicHskHt = PicHsk.Height
  wPicHskWdt = PicHsk.Width
  
  wPicDPHt = PicDP.Height: wPicDPWdt = PicDP.Width  '*** Jay 3.4[DP]
  
  wFraPic3DHt = FraPic3D.Height
  wFraPic3DWdt = FraPic3D.Width
  wFraPicHskHt = FraPicHsk.Height
  wFraPicHskWdt = FraPicHsk.Width
 
  wFraPic3DLft = FraPic3D.Left
  wFraPic3DTop = FraPic3D.Top
  wFraPicHskLft = FraPicHsk.Left
  wFraPicHskTop = FraPicHsk.Top
  
  wPic3DDblClk = False
  wPicHskDblClk = False
  
  wPicDsgZoomHt = PicDsgZoom.Height
  wPicDsgZoomWdt = PicDsgZoom.Width
  wFraPicDsgZoomHt = FraPicDsgZoom.Height
  wFraPicDsgZoomWdt = FraPicDsgZoom.Width
  wFraPicDsgZoomLft = FraPicDsgZoom.Left
  wFraPicDsgZoomTop = FraPicDsgZoom.Top
  
  
  
   '****** Sachin 3.02.0 - Hiding the Partition Key Column based on Partition Selected
  ADC("DmPrtKey").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdDsgRm.ColProp("DrPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdDsgLab.ColProp("DlPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdDsgAna.ColProp("DaPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdDsgCT.ColProp("DctPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  GrdPrcPrd.ColProp("DppPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
     
  FraDsgMatrix.BackColor = vbWhite      '4.3.0.0
  
  Call SetStockSizeVisible(3, GrdDsgRm.ColProp("DRLN3"), ADC.Connection)
End Sub

Private Sub ADC_KeyWhen()
  '*** DmTcTyp field will get its value from the menu code
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll
  '*** Reset the Picture controls to No Picture
  '*** Restore the original height, width & position of the picture frames

  ADC("DmTcTyp") = UCase(ADC.MenuCd)
  
  Call EnaDisaCmds(True)
  '*** Alternative Code
  'CmdCpy.Enabled = False: CmdSumm.Enabled = False
  
  Call HideAllFras
  
  Call ShowPic(0, 0, "", PicHsk)
  Call ShowPic(0, 0, "", Pic3D)
  
  '3.11.0
  Call ShowPic(0, 0, "", PicDsg1)
  Call ShowPic(0, 0, "", PicDsg2)
  Call ShowPic(0, 0, "", PicDsg3)

  FraPic3D.Height = wFraPic3DHt
  FraPic3D.Width = wFraPic3DWdt
  FraPicHsk.Height = wFraPicHskHt
  FraPicHsk.Width = wFraPicHskWdt
    
  FraPic3D.Left = wFraPic3DLft
  FraPic3D.Top = wFraPic3DTop
  FraPicHsk.Left = wFraPicHskLft
  FraPicHsk.Top = wFraPicHskTop
  
  FraPicDsgZoom.Height = wFraPicDsgZoomHt
  FraPicDsgZoom.Width = wFraPicDsgZoomWdt
  FraPicDsgZoom.Left = wFraPicDsgZoomLft
  FraPicDsgZoom.Top = wFraPicDsgZoomTop
  

  
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "N" Then
    GrdDsgRm.ColProp("DrRmPtr").Style = fgcinvisible
  ElseIf moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then
    GrdDsgRm.ColProp("DrRmPtr").Style = fgcNormal
  End If
    
  'If moCn.GetFldVal("Select HShowPrdRmDets from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'") = "N" Then
  If moCn.GetFldVal("Select HShowPrdRmDets from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "N" Then
    GrdDsgRm.ColProp("DrQty").Style = fgcinvisible
    GrdDsgRm.ColProp("DrWt").Style = fgcinvisible
    lblShBothWts.Visible = False
    ADC("wShowPrdRmDets").Enabled = False
    ADC("wShowPrdRmDets").Visible = False
    wShBothWts = "N"
  'ElseIf moCn.GetFldVal("Select HShowPrdRmDets from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then
  ElseIf moCn.GetFldVal("Select HShowPrdRmDets from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then
    GrdDsgRm.ColProp("DrQty").Style = fgcNormal
    GrdDsgRm.ColProp("DrWt").Style = fgcNormal
    lblShBothWts.Visible = True
    ADC("wShowPrdRmDets").Enabled = True
    ADC("wShowPrdRmDets").Visible = True
    wShBothWts = "Y"
  End If
' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################
  TxtPwd.text = ""
' ###########################################  Manoj  ###########################################

  CmdPushCurr.Enabled = False       '****** Sachin 3.02 [26/11/07] - Purge Partition
  
End Sub

Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
 '*** Jay 2.13(CT) ***
 '*** Force the user to enter Component Code when the Design Type is 'PM' (Part Master)
 '  If ADC("DmTcTyp") = "PM" And Trim(ADC("DmSz")) = "" Then
 '    Cancel = True: ErrMsg = "Enter Component Code": Exit Sub
 '  End If
 '*** Jay 2.13(CT) ***
              

 '****** Sachin 3.02.0 [01/12/07] - Purge Partition Check whether Design exists in Previous Partition
 If gs_Partition = ctCurrPrtn And Not moCn.RecSeek("Select 'x' From DsgMst Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmPrtKey='" + ctCurrPrtn + "'") Then
    If moCn.RecSeek("Select 'x' From DsgMst Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmPrtKey<>'" + ctCurrPrtn + "'") Then _
        Cancel = True: ErrMsg = "Design Exists In Previous Partition, Log On to Previous Partition to View it.": Exit Sub
 End If

  Dim ws_CoCd() As String, i As Integer, ws_ValidUsrForEdit As Boolean, ws_DmValidCoCd As String
  ws_ValidUsrForEdit = True
  ws_DmValidCoCd = ""
  ws_DmValidCoCd = moCn.GetFldVal(" Select DmValidCoCd from DsgMst " + _
                  " Where DmTcTyp ='" + ADC("DmTcTyp") + "'" + _
                  " And DMCd = '" + ADC("DMCd") + "'" + _
                  " And DMSz ='" + ADC("DMSz") + "' And DmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DmPrtKey End)")
  If ws_DmValidCoCd <> "" And ws_MultiCoMod Then
        ws_CoCd = Split(ws_DmValidCoCd, ",")
        If IsArray(ws_CoCd) Then
          For i = 0 To UBound(ws_CoCd)
'            Cancel = Not moCn.RecSeek("Select 1 From Head where " + _
                                      "HCd = 'ZSELF' and HCoCd = '" + Replace$(Replace$(ws_CoCd(i), "'", "''"), """", """""") + "'")
'            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Company Code": Exit Sub
            
            If moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + ws_CoCd(i) + "'+',%')") Then
                ws_ValidUsrForEdit = True
                Exit For
            Else
                ws_ValidUsrForEdit = False
            End If
                        
          Next i
          If Not ws_ValidUsrForEdit Then
                Cancel = True
                 ErrMsg = "Design already exists, you cannot add again and you don't have permissions to view the details"
'                ErrMsg = "User Doesn't Have Access To This Design"
'                ADC.AllowClose = True
'                ADC.AllowDelete = False
'                ADC.AllowSave = True
                
                Exit Sub
            End If
        End If
  End If
 
' errmsg = Design already exists, u cannot add again and u don't have permissions to view the details
 '****** Sachin 3.02.0 [01/12/07] - Purge Partition Check whether Design exists in Previous Partition
 mb_Upd = False
End Sub

Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Enable all the command buttons
  '*** Set the default value of Design Parts as 1, DMDsgDt as '01/01/80' and also
      'set the default values for the fields DmPrdSeq and DmVaCtg from the Head file
  '*** Show the Sketch and 3D pictures based on the Design Type
  '*** In case of 'SM' show only the Sketch Picture
  '*** Disable all the tab frames except the current tab frame of the Tab Page
  
  Dim i As Single
'  Dim ws_CoCd() As String, i As Integer, ws_ValidUsrForEdit As Boolean
'  ws_ValidUsrForEdit = True
'  If adc.AddFndMode = xFndMode And ws_MultiCoMod Then
'        ws_CoCd = Split(adc("DmValidCoCd"), ",")
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
'
'          Next i
'          If Not ws_ValidUsrForEdit Then
'                Cancel = True
'                ErrMsg = "User Doesn't Have Access To This Design"
''                ADC.AllowClose = True
'                adc.AllowDelete = False
'                adc.AllowSave = True
'
'                Exit Sub
'            End If
'        End If
'  End If
'
  'KBS.134-Special characters not allowed for Design code
  If ADC.AddFndMode = xaddmode Then
    For i = 1 To Len(ADC("DmCd"))
      If InStr(1, " _0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-.", Mid(ADC("DmCd"), i, 1)) < 1 Then _
         Cancel = True: ErrMsg = "Special Characters Should Not Be Allowed": Exit Sub
    Next
  End If
  Call EnaDisaCmds(False)
  '*** Alternative Code
  'FraNKeyAll.Enabled = True
  'FraCpy.Enabled = False
  'FraSumm.Enabled = False
  'CmdCpy.Enabled = True
  'CmdSumm.Enabled = True
  
  
  Call GetLockDt(gs_CoCd, "DSGMST", ADC.MenuCd, ms_LockDt, ms_FutureDt)
  
  mf_DtLocked = IIF((ADC("DmCreatedDt") <> "" And ADC("DmCreatedDt") < CDate(ms_LockDt)) _
                                Or _
                   (ADC("DmCreatedDt") <> "" And ADC("DmCreatedDt") > CDate(ms_FutureDt)) _
                     , True, False)
   
  
  '****** Sachin 3.02.0 - [01/12/07] - Purge Partition - Setting the Default Value of Prt Key to 'C'
  If ADC.IsNew = True And Trim(ADC("DmPrtKey")) = "" Then ADC("DmPrtKey") = ctCurrPrtn
  CmdPushCurr.Enabled = IIF(ADC("DmSz") = "" And ADC("DmPrtKey") <> ctCurrPrtn, True, False)
  '****** Sachin 3.02 - Setting the Created Date From Current Server Date
  '6.1
  If ADC.IsNew = True Then ADC("DmCreatedDt") = Format$(moCn.SrvrDate, DtShortStrWithDash()): ADC("DMValidYN") = "Y"
  '****** Sachin 3.02 [26/11/07] - Purge Partition
  
  If ADC("DmParts") = 0 Then ADC("DmParts") = 1
  wHskPth = GetPictPath(ADC("DmTcTyp"), True, ADC("DmCd"), ADC("DmSz"))
  w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"))
  
  ' **** Zubin 211 **** '
  'wHskPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), True)
  'w3DPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), False)
  ' **** Zubin 211 **** '
  
  Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
  Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
  
  'Sachin 3.11.0 - 3 New Images in Design Master
  w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "1")
  Call ShowPic(2955, 2955, w3DPth, PicDsg1) '6.1_24 Ht, wt changed to avoid pic distort
  w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "2")
  Call ShowPic(2955, 2955, w3DPth, PicDsg2) '6.1_24 Ht, wt changed to avoid pic distort
  w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "3")
  Call ShowPic(2955, 2955, w3DPth, PicDsg3) '6.1_24 Ht, wt changed to avoid pic distort
  '****uma
  'If ADC("DmPrdSeq") = "" And (ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM") Then _
    ADC("DmPrdSeq") = moCn.GetFldVal("Select hDmPrdSeq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
      
  If ADC("DmVaCtg") = "" Then ADC("DmVaCtg") = moCn.GetFldVal("Select hDmVaCtg from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
  'If ADC("DMDsgDt") = "" Then ADC("DMDsgDt") = MWLib.EmptyDate     '*** Jay 3.01 *** [DMDsgDt= SrvrDt]
  If ADC("DMDsgDt") = "" Then ADC("DMDsgDt") = moCn.SrvrDate        '*** Jay 3.01 *** [DMDsgDt= SrvrDt]
' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################
  ms_PwdFlag = False
  If ADC("DMLockYN") = "" Then ADC("DMLockYN") = "N": ms_PwdFlag = True
  If ADC("DMLockYN") = "Y" Then
    TxtPwd.Enabled = True
    ms_PwdFlag = False
    '*************Geeta*******************
    ADC.AllowDelete = False
    '*************************************
  Else
    TxtPwd.Enabled = False
    ms_PwdFlag = True
    '*************Geeta*******************
    ADC.AllowDelete = True
    '************************************
  End If
  '*** Jay 2.14 *** (SEO Change)
  'ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  'ms_Pwd = moCn.GetFldVal("Select PValue1 from Param where " + _
                    "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'")
  'If ms_Pwd = "" Then ms_PwdFlag = True
  '*** Jay 2.14 *** (SEO Change)
' ###########################################  Manoj  ###########################################
  For i = 0 To TabDesgnDet.Tabs - 1
    FraDesgnDet(i).Enabled = False
  Next i
  FraDesgnDet(TabDesgnDet.Tab).Enabled = True
'*** Jay 2.13(CT) ***
  If ADC.MenuCd = "DM" And ADC("DMSZ").Value = "" Then
      TabDesgnDet.TabVisible(6) = True: TabDesgnDet.TabEnabled(6) = True
      TabDesgnDet.TabVisible(10) = True: TabDesgnDet.TabEnabled(10) = True   '*** Bhavna
      If wModAvl = "Y" Then
        TabDesgnDet.TabVisible(7) = True: TabDesgnDet.TabEnabled(7) = True
        ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
        If mb_DsgCat = True Then
          TabDesgnDet.TabVisible(8) = True: TabDesgnDet.TabEnabled(8) = True
          TabDesgnDet.TabsPerRow = 12   '578-12 changed 11 to 12        '*** Jay 3.4[DP]
        Else
          TabDesgnDet.TabVisible(8) = False: TabDesgnDet.TabEnabled(8) = False
          TabDesgnDet.TabsPerRow = 11   '578-12 changed 10 to 11
        End If
        ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
      ElseIf wModAvl = "N" Then
        TabDesgnDet.TabVisible(7) = False: TabDesgnDet.TabEnabled(7) = False
        ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
        If mb_DsgCat = True Then
          TabDesgnDet.TabVisible(8) = True: TabDesgnDet.TabEnabled(8) = True
          TabDesgnDet.TabsPerRow = 11   '578-12 changed 10 to 11        '*** Jay 3.4[DP]
        Else
          TabDesgnDet.TabVisible(8) = False: TabDesgnDet.TabEnabled(8) = False
          TabDesgnDet.TabsPerRow = 10    '578-12 changed 9 to 10
        End If
        ' ******* Manali 3.4.1 - 22/10/08 - Design Catalogue Moduele
      End If
      
      If wb_CTMod = False Then
        GrdDsgCT.ColProp("DctChr").Style = fgcinvisible
        GrdDsgCT.ColProp("DctRmCd").Style = fgcinvisible
        GrdDsgCT.ColProp("DctDmCol").Style = fgcinvisible
        GrdDsgCT.ColProp("DctRmWt").Style = fgcinvisible
        GrdDsgCT.ColProp("DctDc").Style = fgcinvisible
        GrdDsgCT.ColProp("DctPrdSeq").Style = fgcinvisible
        GrdDsgCT.ColProp("DctDesc").Style = fgcinvisible
        GrdDsgCT.Width = 5000
      Else
        GrdDsgCT.ColProp("DctChr").Style = fgcNormal
        GrdDsgCT.ColProp("DctRmCd").Style = fgcNormal
        GrdDsgCT.ColProp("DctDmCol").Style = fgcNormal
        GrdDsgCT.ColProp("DctRmWt").Style = fgcNormal
        GrdDsgCT.ColProp("DctDc").Style = fgcNormal
        GrdDsgCT.ColProp("DctPrdSeq").Style = fgcNormal
        GrdDsgCT.ColProp("DctDesc").Style = fgcNormal
      End If
      
  ElseIf ADC.MenuCd = "DM" And ADC("DMSZ").Value <> "" Then
      TabDesgnDet.TabVisible(6) = False: TabDesgnDet.TabEnabled(6) = False
      TabDesgnDet.TabVisible(7) = False: TabDesgnDet.TabEnabled(7) = False
      TabDesgnDet.TabVisible(8) = False: TabDesgnDet.TabEnabled(8) = False   '*** Jay 3.4[DP]
      TabDesgnDet.TabVisible(10) = False: TabDesgnDet.TabEnabled(10) = False   '*** bhavna
      TabDesgnDet.TabsPerRow = 8    '578-12 changed 7 to 8
      
  End If
  
  '***** Sachin - Disabling the Delete Button for Previous Partition Designs
  ADC.AllowDelete = IIF(ADC("DmPrtKey") = ctCurrPrtn, True, False)
  
  '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSize Master]******
  mb_VldRmLn = moCn.RecSeek("Select 'x' from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' And hChkVldRmLnYN='Y'")
  '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSize Master]******
  
  'vm.2- default setting for Rm change Grid and check box STW
  ADC("wCpyFgBagGrp") = "-"
  GrdCpyFgBag.ColProp("DtCol").Style = fgcinvisible
  GrdCpyFgBag.ColProp("DtRmRSz").Style = fgcinvisible
  ChkFCSTW.Value = vbUnchecked
'  '****** Sachin 2.13.0 - 02-06-2006 ******
'  If ADC.MenuCd = "DM" Then
'    If ADC("DMSZ").Value = "" And wModAvl = "Y" Then
'      TabDesgnDet.TabVisible(7) = True: TabDesgnDet.TabEnabled(7) = True
'      TabDesgnDet.TabsPerRow = 8
'    Else
'      TabDesgnDet.TabVisible(7) = False: TabDesgnDet.TabEnabled(7) = False
'      TabDesgnDet.TabsPerRow = 7
'    End If
'  End If
'  '****** Sachin 2.13.0 - 02-06-2006 ******
'*** Jay 2.13(CT) ***
'If mf_DtLocked Then
'
'End If


End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, _
                          Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** For DmSz show help of Size if Design Type is 'DM' or 'SM' and show help of Part No
      'if Design Type is 'PM'
      'Similarly for the working field wDmSz (In the copy frame)
  '*** DmHldDesc
      'If DmHld= 'N' then cannot change Holding Description
  '*** DmCol
      '(Old Code) If Record Exists in DsgRm then Cannot Change Design Colour
'*** Jay 2.14 *** (SEO Change)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################
'  If UCase(IdName) <> UCase("DmCd") And UCase(IdName) <> UCase("DmSz") _
    And UCase(IdName) <> UCase("wShowPrdRmDets") And UCase(IdName) <> UCase("DmLockYN") _
    And ADC.AddFndMode = xFndMode And ms_PwdFlag = False Then _
      Cancel = True: ErrMsg = "Design is Locked, Open Lock to Add/Edit/Delete the record": Exit Sub
'  If UCase(IdName) <> UCase("DmCd") And UCase(IdName) <> UCase("DmSz") _
'     And UCase(IdName) <> UCase("wShowPrdRmDets") And ms_PwdFlag = False _
'     And ADC.AddFndMode = xFndMode Then Cancel = True: ErrMsg = "Enter SEO Password To Edit": Exit Sub
' ###########################################  Manoj  ###########################################

  '****** Sachin 3.02.0 - Purge Partition - Previous Partition Records to be Locked
  If ADC.AddFndMode = xFndMode And ADC("DmPrtKey") <> ctCurrPrtn _
      And UCase(IdName) <> UCase("DmCd") And UCase(IdName) <> UCase("DmSz") _
      And UCase(IdName) <> UCase("wShowPrdRmDets") And UCase(IdName) <> UCase("wDmCtg") Then _
   Cancel = True: ErrMsg = "Design In Previous Partition Cannot Be Edited ": Exit Sub

  If UCase(IdName) <> UCase("DmCd") And UCase(IdName) <> UCase("wDmCtg") And UCase(IdName) <> UCase("DmSz") _
    And UCase(IdName) <> UCase("wShowPrdRmDets") And UCase(IdName) <> UCase("DmLockYN") _
    And ADC.AddFndMode = xFndMode And ADC("DmLockYN") = "Y" Then _
      Cancel = True: ErrMsg = "Design is Locked, Open Lock to Add/Edit/Delete the record": Exit Sub
  '*** Jay 2.14 *** (SEO Change)
    
  Select Case UCase(IdName)
  Case Is = UCase("DmCd")
   ' Call HlpList.DmCd(IIF(ADC("DmTcTyp") = "PM", "DM", ADC("DmTcTyp")))
    Call HlpList.DmCd(ADC("DmTcTyp"), ADC("wDmCtg"))
  '*** Jay 3.2.0 [DefSz]
  Case Is = UCase("DmSz"), UCase("DmDefSz")
    If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then
      Call HlpList.PMCd("DMSZ")
    ElseIf ADC("DmTcTyp") = "PM" Then
      If UCase(IdName) = UCase("DmDefSz") Then Cancel = True: ErrMsg = "Cannot Define": Exit Sub
      Call HlpList.DmSz("PM", ADC("DmCd"))
    End If
  '*** Jay 3.2.0 [DefSz]
  Case Is = UCase("DmCtg")
    If ADC.AddFndMode = xaddmode Then ADC("DmCtg") = ADC("wDmCtg")
      Call HlpList.PMCd("DMCTG")
'      Cancel = True
  Case Is = UCase("DmValidCoCd")
    HlpList.MultiSelect = True
    Call HlpList.hCoCd
 
 Case Is = UCase("wDmCtg")
      Call HlpList.PMCd("DMCTG")
'' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################
  Case Is = UCase("DmLockYN")
    '*** Jay 2.14 *** (SEO Change)
    If ADC("DmLockYN") = "Y" And ms_PwdFlag = False And Trim$(TxtPwd.text) = "" And moCn.RecSeek("Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1<>'' ") Then _
                       Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
    'If Trim$(TxtPwd.text) <> Trim$(ms_Pwd) And ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to Edit Value": Exit Sub
    Call HlpList.PMCd("YN")
    '*** Jay 2.14 *** (SEO Change)
' ###########################################  Manoj  ###########################################
  Case Is = UCase("DMPrdCtg")
      Call HlpList.PMCd("PRDCTG")
  'uni.4 help for SalCtg 2 and 3
  Case Is = UCase("DmSalCtg"), UCase("DmSalCtg2"), UCase("DmSalCtg3")
  'Geeta*****Crm***2.12
      Call HlpList.PSCd("SALCTG", ADC("DmCtg"))
      
  Case Is = UCase("DmCmCd")
      HlpList.MultiSelect = True      '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Exclusivity]
      Call HlpList.CustCd("C")
  Case Is = UCase("DmPrfVendCd")
      '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Exclusivity]
      Call HlpList.CustCd("S")
      
  Case Is = UCase("DmHld")
      Call HlpList.PMCd("YN")
  Case Is = UCase("DmHldDesc")
      If Trim(ADC("DmHld")) = "" Or ADC("DmHld") = "N" Then _
         Cancel = True: ErrMsg = "No Holding Desc Should Be Entered": Exit Sub
  Case Is = UCase("DmCol")
      '###
      'If GrdDsgRm.Rows - 1 > 0 Then _
         Cancel = True: ErrMsg = "Cannot Change Colour When Rm Records Exist": Exit Sub
      '###
      Call HlpList.PMCd("DMCOL")
  '****uma
  Case Is = UCase("DmPrcsSeq")
  '    HlpList.MultiSelect = True  'EmrSv
      Call HlpList.vPMCd(gs_CoCd, "PRCSSEQ")
  Case Is = UCase("DmPrdSeq")
'      HlpList.MultiSelect = True  'EmrSv
      Call HlpList.vPMCd(gs_CoCd, "PRDSEQ")
  Case Is = UCase("DmUom")
      Call HlpList.PMCd("DMUOM")
  Case Is = UCase("DmVaCtg")
    Call HlpList.PMCd("VACTG")
  Case Is = UCase("DmKt")
    Call HlpList.PMCd("KT")
  Case Is = UCase("wDmTcTyp")
    Call HlpList.PMCd("DMTCTYP")
  Case Is = UCase("wDmCd")
    Call HlpList.DmCd(ADC("wDmTcTyp"))
  Case Is = UCase("wDmSz")
    If ADC("wDmTcTyp") = "DM" Or ADC("wDmTcTyp") = "SM" Then
      Call HlpList.PMCd("DMSZ")
    ElseIf ADC("wDmTcTyp") = "PM" Then
      Call HlpList.DmSz("PM", ADC("wDmCd"))
    End If
  Case Is = UCase("DmLsCtg")
    Call HlpList.PMCd("LSCTG")
  Case Is = UCase("DmSetCd")
    Call HlpList.DmSetCd
  Case Is = UCase("DmRegnExc")
    Call HlpList.PMCd("YN")
  Case Is = UCase("DmCmGrp")
    Call HlpList.PMCd("YN")
  Case Is = UCase("wCmName")
    Call HlpList.CmName("C")
  Case Is = UCase("wShowPrdRmDets")
    Call HlpList.PMCd("YN")
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  Case Is = UCase("wCoCd")
    Call HlpList.hCoCd
  Case Is = UCase("wBYy")
    Call HlpList.PMCd("Yy")
  Case Is = UCase("wBChr")
    Call HlpList.PMCd("BCHR", "'P'")
  Case Is = UCase("wBNo")
    'vm.2 if copy from fg bag then calling help of bag no, if from master record calling help of stkno
    If OptFgbag.Value = True Then
      Call HlpList.BNo(ADC("wCoCd"), ADC("wBYy"), ADC("wBChr"), "N", "PFG", , , , , ADC("DmCd"), True)
    ElseIf OptMstRec.Value = True Then
      Call HlpList.OdDmStkNo(gs_CoCd, ADC("DmCd"), "", ADC("DmSz"), ADC("wCoCd"), ADC("wBYy"), ADC("wBChr"))
    End If
  Case Is = UCase("DmValidYn")     '4.3.0.0
    Call HlpList.PMCd("YN")
  'vm.2 calling help for Rm group and Rm category
  Case Is = UCase("wCpyFgBagGrp")
    Call HlpList.PMCd("ORDCPYGR")
  Case Is = UCase("wCpyFgBagCtg")
    Call HlpList.PMCd("GLDAS")
  
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  End Select
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** DmHldDesc on DmHld         Normal
      'If DmHld='N' or '' then ''
  '*** DmBagPcs on DmCtg          Normal
      'PNum('DMCTG', DmCtg, '')
  '*** wCmRegnCd on DmCmCd        Normal, Init
      'CmRegnCd from CustMst(CmCtg= 'C', CmCd= DmCmCd)
  '*** DmRegnExc on DmCmCd        Normal
      'If DmCmCd= Blank then DmRegnExc= Blank
  '*** DmCmCd on wCmName          Normal
      'If not the same as the existing value then DmCmCd= CmCd from CustMst(CmCtg= 'C', CmName= wCmName)
  '*** wCmName on DmCmCd          Normal
      'If not the same as the existing value then wCmName= CmName from CustMst(CmCtg= 'C', CmCd= DmCmCd)
  
  Dim wCmCd As String, wCustName As String
  Select Case UCase(IdName)
  Case Is = UCase("DmHldDesc")
    If ADC.Mode = xNorm Then
      If ADC("DmHld") = "N" Or Trim(ADC("DmHld")) = "" Then ADC("DmHldDesc") = ""
    End If
    
        Case Is = UCase("DmBagPcs")
    If ADC.Mode = xNorm Then
      If SourceName <> "WDMCTG" Then
        ADC("DmBagPcs") = moCn.GetFldVal("Select PNum From Param " + _
                            "where PTyp='DMCTG' and PMCd='" + ADC("DmCtg") + "'")
      End If
    End If
  Case Is = UCase("wCmRegnCd")
  '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******
  Dim wCustRegn() As String
  wCustRegn = Split(ADC("DmCmCd"), ",")
    If (ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy) And Trim(ADC("DmCmCd")) <> "" Then _
       ADC("wCmRegnCd") = moCn.GetFldVal("Select CmRegnCd From CustMst " + _
                          "where CmCtg= 'C' and CmCd='" + wCustRegn(0) + "'")
  '****** Sachin 3.3.0 (27-06-08) - [37. Design Exclusivity for Multiple Customers] ******
  
  Case Is = UCase("DmRegnExc")
  
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
       If ADC("DmCmCd") = "" Then
          ADC("DMREGNEXC").Enabled = False: ADC("DmRegnExc") = ""
       Else
          ADC("DmRegnExc").Enabled = Not MultiRegnCust(ADC("DmCmCd"))
          If MultiRegnCust(ADC("DmCmCd")) = True Then ADC("DmRegnExc") = "N"
       End If
     End If
  
  Case Is = UCase("DmCmGrp")
  
    If ADC.Mode = xNorm Or ADC.Mode = xInit Or ADC.Mode = xCopy Then
       If ADC("DmCmCd") = "" Then
          ADC("DMCMGRP").Enabled = False: ADC("DmCmGrp") = ""
       Else
          ADC("DmCmGrp").Enabled = Not MultiRegnCust(ADC("DmCmCd"))
          If MultiRegnCust(ADC("DmCmCd")) = True Then ADC("DmCmGrp") = "N"
       End If
     End If
  '  Case Is = UCase("DmCmCd")
  '    If ADC.Mode = xNorm Then
  '      wCmCd = moCn.GetFldVal("Select CmCd from CustMst where CmCtg='C' and " + _
  '                             "CmName= '" + ADC("wCmName") + "'")
  '      If wCmCd <> ADC("DmCmCd") Then ADC("DmCmCd") = wCmCd
  '    End If
  '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Region Exclusivity] ******
  
  Case Is = UCase("wCmName")
    If ADC.Mode = xNorm Then
       wCustName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
                                  "CmCd='" + ADC("DmCmCd") + "'")
       If wCustName <> ADC("wCmName") Then ADC("wCmName") = wCustName
    ElseIf ADC.Mode = xInit Then
       ADC("wCmName") = moCn.GetFldVal("Select CmName from CustMst where CmCtg='C' and " + _
                                  "CmCd='" + ADC("DmCmCd") + "'")
    End If
  ' Zubin 212
 Case Is = UCase("DmUOM")
    If ADC.Mode = xNorm Then
      If SourceName <> "WDMCTG" Then
        ADC("DmUOM") = moCn.GetFldVal("Select PValue From Param " + _
                        "where PTyp='DMCTG' and PMCd='" + ADC("DmCtg") + "'")
      End If
    End If
  ' Zubin 212
  
  'EmrSv
  Case Is = UCase("DmPrcsSeq")
    If ADC.Mode = xNorm Then
      Dim wReq As String
      If ADC("DmPrcsSeq") = "" Then
        wReq = moCn.GetFldVal("Select PDesc225 from Param where PTyp='DMCTG' and PMCD='" + ADC("DMCTG") + "' and PSCd=''")
        If wReq = "" Then wReq = moCn.GetFldVal("Select hDmPrcsSeq from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")
        ADC("DmPrcsSeq") = wReq
      End If
    End If
  End Select
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** DmCd         Not Blank
  '*** DmSz         If DmTcTyp='DM' or 'SM' then valid PMCd('DMSZ') or ''
  '                 If DmTcTyp='PM' then not blank
  '*** DmCtg        If DmTcTyp='DM' or 'SM' then valid PMCd('DMCTG')
  '                 If DmTcTyp='PM' then '' or valid PMCd('DMCTG')
  '*** DmPrdCtg     If DmTcTyp='DM' or 'SM' then valid PMCd('PRDCTG')
  '                 If DmTcTyp='PM' then '' or valid PMCd('PRDCTG')
  '*** DmSalCtg     If DmTcTyp='DM' or 'SM' then valid PMCd('SALCTG')
  '                 If DmTcTyp='PM' then '' or valid PMCd('SALCTG')
  '*** DmBagPcs     If DmTcTyp='DM' or 'SM' then >0
  '*** DmParts      >0
  '*** DmCmCd       '' or valid CmCd from CustMst(CmCtg= 'C')
  '*** DmHld        valid PMCd('YN')
  '*** DmCol        valid PMCd('DMCOL')
  '*** DmPrdSeq     If DmTcTyp='DM' or 'SM' then valid PMCd('PRDSEQ')
  '                 If DmTcTyp='PM' then '' or valid PMCd('PRDSEQ')
  '*** DmUOM        If DmTcTyp='DM' or 'SM' then valid PMCd('DMUOM')
  '                 If DmTcTyp='PM' then '' or valid PMCd('DMUOM')
  '*** DmVaCtg      If DmTcTyp='DM' or 'SM' then valid PMCd('VACTG')
  '                 If DmTcTyp='PM' then '' or valid PMCd('VACTG')
  '*** DmKt         valid PMCd('KT')
  '*** DMDsgDt      valid Date
  '*** DmRegnExc    '' or valid PMCd('YN')
  '*** wCmName      '' or valid CmName from CustMst(CmCtg= 'C')
  ' **DmCmGrp       '' or valid PMCd('YN')
 '6.1b-46 mb_Upd check removed, it will check when Rm,lab record change
 'If ADC(IdName) <> pv_NewValue And UCase(IdName) <> UCase("DMCD") And UCase(IdName) <> UCase("DMSZ") And UCase(IdName) <> UCase("DMTCTYP") And UCase(IdName) <> UCase("DmLockYN") Then mb_Upd = True
 Dim ws_CoCd() As String, i As Integer, ws_PrcsSeqArr() As String
  Select Case UCase(IdName)
     Case Is = UCase("DmCd")
       If Trim(pv_NewValue) = "" Then
         Cancel = True: ErrMsg = "Design Code Cannot Be Blank": Exit Sub
       End If
     '*** Check for Part Code when TcTyp= 'PM' has been done in KeyValid itself
     Case Is = UCase("DmSz")
       If pv_NewValue <> "" And _
         (ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM") Then
         Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                              " PTyp= 'DMSZ' and PMCd = '" + pv_NewValue + "'")
         ErrMsg = "Invalid Size": Exit Sub
       End If
     '*** Jay 3.2.0 [DefSz]
     Case Is = UCase("DmDefSz")
        If pv_NewValue <> "" And (ADC("DmSz") <> "" Or ADC("DmTcTyp") = "PM") Then _
            Cancel = True: ErrMsg = "Cannot Enter Default Size": Exit Sub
        If pv_NewValue <> "" And (ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM") Then
            Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                            " PTyp= 'DMSZ' and PMCd = '" + pv_NewValue + "' and PValidYn = 'Y'")
            ErrMsg = "Invalid Size/ Size Not Defined": Exit Sub
        End If
     '*** Jay 3.2.0 [DefSz]
    Case Is = UCase("DmCtg")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Category Cannot Be Blank": Exit Sub
      Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'DMCTG' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Category":  Exit Sub
       
      End If
        
    Case Is = UCase("DmValidCoCd")
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
  
    Case Is = UCase("wDmCtg")
'      If pv_NewValue = "" Then
'        If adc("DmTcTyp") = "DM" Or adc("DmTcTyp") = "SM" Then _
'           Cancel = True: ErrMsg = "Category Cannot Be Blank": Exit Sub
'      Else
        Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'DMCTG' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Category": Exit Sub
'      End If
  ' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################
    Case Is = UCase("DmLockYN")
      If pv_NewValue = "" Then
        Cancel = True: ErrMsg = "Design Lock (Y/N) Cannot Be Blank"
      Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Design Lock (Y/N)"
        If Cancel = False Then
          If pv_NewValue = "Y" Then
            TxtPwd.Enabled = True: ms_PwdFlag = False
             '*************Geeta*******************
              ADC.AllowDelete = False
             '************************************
          Else
            If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                     "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                     "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                     "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
              TxtPwd.text = "": TxtPwd.Enabled = False: ms_PwdFlag = True
              '*************Geeta*******************
              ADC.AllowDelete = True
              '************************************
            End If
          End If
        End If
      End If
      Exit Sub
  ' ###########################################  Manoj  ###########################################
  
    Case Is = UCase("DmPrdCtg")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Production Category Cannot Be Blank": Exit Sub
      Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'PRDCTG' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Production Category": Exit Sub
      End If
    Case Is = UCase("DmSalCtg")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Sales Category Cannot Be Blank": Exit Sub
      Else
        'uni.4 value of SalCtg,SalCtg 2 and 3 should no be same, except '-'
        If pv_NewValue <> "-" Then
          If pv_NewValue = ADC("DmSalCtg2") Or pv_NewValue = ADC("DmSalCtg3") Then _
            Cancel = True: ErrMsg = "Sales Categories Should Not Be Same": Exit Sub
        End If
        '******Geeta****Crm***2.12****PMCd=DmCtg and PScd=DmSalCtg
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'SALCTG' and PMCd = '" + ADC("DmCtg") + "' and PSCd= '" + pv_NewValue + "'")
        ErrMsg = "Invalid Sales Category": Exit Sub
      End If
    'uni.4 validation for SalCtg2
    Case Is = UCase("DmSalCtg2")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Sales Category2 Cannot Be Blank": Exit Sub
      Else
        'uni.4 value of SalCtg,SalCtg 2 and 3 should no be same, except '-'
        If pv_NewValue <> "-" Then
          If pv_NewValue = ADC("DmSalCtg") Or pv_NewValue = ADC("DmSalCtg3") Then _
            Cancel = True: ErrMsg = "Sales Categories Should Not Be Same": Exit Sub
        End If
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'SALCTG' and PMCd = '" + ADC("DmCtg") + "' and PSCd= '" + pv_NewValue + "'")
        ErrMsg = "Invalid Sales Category2": Exit Sub
      End If
    'uni.4 validation for SalCtg3
    Case Is = UCase("DmSalCtg3")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Sales Category3 Cannot Be Blank": Exit Sub
      Else
        'uni.4 value of SalCtg,SalCtg 2 and 3 should no be same, except '-'
        If pv_NewValue <> "-" Then
          If pv_NewValue = ADC("DmSalCtg") Or pv_NewValue = ADC("DmSalCtg2") Then _
            Cancel = True: ErrMsg = "Sales Categories Should Not Be Same": Exit Sub
        End If
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'SALCTG' and PMCd = '" + ADC("DmCtg") + "' and PSCd= '" + pv_NewValue + "'")
        ErrMsg = "Invalid Sales Category3": Exit Sub
      End If
    Case Is = UCase("DmBagPcs")
      If pv_NewValue <= 0 Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Bag Opening Pieces Should Be Greater Than Zero": Exit Sub
      End If
    Case Is = UCase("DmParts")
      If pv_NewValue <= 0 Then
        Cancel = True: ErrMsg = "No Of Parts Should Be Greater Than Zero": Exit Sub
      End If
      
    Case Is = UCase("DmCmCd")
    '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Region Exclusivity] ******
    If pv_NewValue <> "" Then
          Dim ws_CstArr() As String, CustCtr As Integer
          ws_CstArr = Split(pv_NewValue, ",")
          If IsArray(ws_CstArr) Then
            If UBound(ws_CstArr) > 8 Then Cancel = True: ErrMsg = "Only 9 Customers Are Allowed for Region Exclusivity ": Exit Sub
            For CustCtr = 0 To UBound(ws_CstArr)
                Cancel = Not moCn.RecSeek(" Select 'x' From CustMst where CmCtg= 'C' and CmCd = '" + ws_CstArr(CustCtr) + "'")
              If Cancel = True Then ErrMsg = "Customer (" + CStr(CustCtr + 1) + ") Is An Invalid Customer": Exit Sub
            Next CustCtr
          End If
     Else
        ADC("WCMREGNCD") = ""
     End If
     
     If MultiRegnCust(pv_NewValue) = True Then ADC("DMREGNEXC") = "N"
     ADC("DMREGNEXC").Enabled = Not MultiRegnCust(pv_NewValue)
     
     If MultiRegnCust(pv_NewValue) = True Then ADC("DMCMGRP") = "N"
     ADC("DMCMGRP").Enabled = Not MultiRegnCust(pv_NewValue)
     
     '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Region Exclusivity] ******
     
     ' ****** Sachin 3.3.0 (27-06-08) - Before 3.3.0 ******
     '      If pv_NewValue <> "" Then
     '        Cancel = Not moCn.RecSeek("select CmCd from CustMst where CmCtg= 'C' " + _
     '                             "and CmCd = '" + pv_NewValue + "'")
     '        ErrMsg = "Invalid Customer Code": Exit Sub
     '      End If
     ' ****** Sachin 3.3.0 (27-06-08) - Before 3.3.0 ******

  Case Is = UCase("DmPrfVendCd")
    '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Region Exclusivity] ******
    If pv_NewValue <> "" Then
            
              Cancel = Not moCn.RecSeek(" Select 'x' From CustMst where CmCtg= 'S' and CmCd = '" + pv_NewValue + "'")
              If Cancel = True Then ErrMsg = "Invalid Preferred Vendor Code": Exit Sub
    End If
    Case Is = UCase("DmHld")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' Or 'N' For Design On Hold": Exit Sub
      
    '4.3.0.0
    Case Is = UCase("DmValidYn")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Enter 'Y' Or 'N' For Design Valid": Exit Sub
      
    Case Is = UCase("DmCol")
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp= 'DMCOL' and PMCd = '" + pv_NewValue + "'")
      '*** Jay 2.13(CT) ***
      If Cancel = True Then ErrMsg = "Invalid Colour": Exit Sub
      If Cancel = False And ADC.MenuCd = "DM" And ADC("DMSZ") = "" And wb_CTMod = True Then
'          Dim i As Integer
          i = 1
          Do While i < GrdDsgCT.Rows
            If GrdDsgCT.Value(i, "DctDmCol") = ADC("DMCOL") Then GrdDsgCT.Value(i, "DctDmCol") = pv_NewValue
            i = i + 1
          Loop
          ADC.AddSave "GrdDsgCt", asAll
      End If
      '*** Jay 2.13(CT) ***
      '**uma
    Case Is = UCase("DmPrcsSeq")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Process Sequence Cannot Be Blank": Exit Sub
      Else
      ws_PrcsSeqArr = Split(pv_NewValue, ",")
        For i = 0 To UBound(ws_PrcsSeqArr)
          If IsArray(ws_PrcsSeqArr) Then
          Cancel = Not moCn.RecSeek("select vPMCd from vParam where vPCoCd = '" + gs_CoCd + "'" + _
                             " and vPTyp= 'PRCSSEQ' and vPMCd = '" + ws_PrcsSeqArr(i) + "'")
            If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Process Sequence": Exit Sub
          End If
        Next i
      'ErrMsg = "Invalid Process Sequence": Exit Sub
      End If
    Case Is = UCase("DmPrdSeq")
      'If pv_NewValue = "" Then
      '  If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
      '     Cancel = True: ErrMsg = "Production Sequence Cannot Be Blank": Exit Sub
      'Else
      ws_PrcsSeqArr = Split(pv_NewValue, ",")
      For i = 0 To UBound(ws_PrcsSeqArr)
        If IsArray(ws_PrcsSeqArr) Then
        Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select vPMCd from vParam where  vPCoCd = '" + gs_CoCd + "'" + _
                             " and vPTyp= 'PRDSEQ' and vPMCd = '" + ws_PrcsSeqArr(i) + "'")
          If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is An Invalid Production Sequence": Exit Sub
        End If
      Next i
        'ErrMsg = "Invalid Production Sequence": Exit Sub
      'End If
    Case Is = UCase("DmUom")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Unit Of Measurement Cannot Be Blank": Exit Sub
      Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'DMUOM' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Unit Of Measurement": Exit Sub
      End If
    Case Is = UCase("DmVaCtg")
      If pv_NewValue = "" Then
        If ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM" Then _
           Cancel = True: ErrMsg = "Value Addition Category Cannot Be Blank": Exit Sub
      Else
        Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'VACTG' and PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Invalid Value Addition Category": Exit Sub
      End If
    Case Is = UCase("DmKt")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'KT' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Karat": Exit Sub
    Case Is = UCase("DMDsgDt")
      If pv_NewValue = "" Then Cancel = True: ErrMsg = "Invalid Design Date": Exit Sub
    Case Is = UCase("DmLsCtg")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'LSCTG' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Loss Category": Exit Sub
    Case Is = UCase("DmRegnExc")
      '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Exclusivity] ******
      Cancel = (pv_NewValue <> "") And MultiRegnCust(ADC("DmCmCd")) = False And (Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Specify Whether The Design Is Exclusive For the Customer(s) In A Region/ All Regions (Enter Y/ N)": Exit Sub
    Case Is = UCase("DmCmGrp")
      '****** Sachin 3.3.0 (27-06-08) - [37. Multiple Customer Exclusivity] ******
      Cancel = (pv_NewValue <> "") And MultiRegnCust(ADC("DmCmCd")) = False And (Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Specify Whether The Design Exclusivity is to be applied to the Customer Group (Enter Y/ N)": Exit Sub
    
    'Case Is = UCase("wCmName")
    '  Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select CmName from CustMst where CmCtg= 'C' and CmName = '" + pv_NewValue + "'"))
    '  If Cancel = True Then ErrMsg = "Invalid Customer Name": Exit Sub
    Case Is = UCase("wShowPrdRmDets")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           "PTyp= 'YN' and PMCd = '" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Enter 'Y' Or 'N' For Prd Or Sales": Exit Sub
       'UCase("wGrsQty"), UCase("wDiaQty"), UCase("wCSQty"), UCase("wAccQty"), _
                UCase ("wGrsWt"), UCase("wGldWt"), UCase("wPlWt"), UCase("wOthMtWt"), _
                UCase("wDiaWt"), UCase("wCSWt"), UCase("wAccWt")
      Call Summary(pv_NewValue)
    '*** Jay 2.14 *** [FG Bag Rm Copy]
    Case Is = UCase("wCoCd")
      Cancel = (pv_NewValue = ctSelfCoCd)
      If Cancel = True Then ErrMsg = "Company Code Cannot be 'ZZZ'": Exit Sub
      Cancel = Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Company Code": Exit Sub
      Cancel = Not moCn.RecSeek("Select 'x' From UsrAccess Where UaMnuCd= '" + ADC.MenuCd + "' And " + _
                         "UaUsrCd = '" + gs_UsrCd + "' And UaMnuForm='MDI' And (UaValidCoCd='' Or ''','+UaValidCoCd+',''' Like '%,'+'" + pv_NewValue + "'+',%')")
      If Cancel = True Then ErrMsg = "User Doesn't Have Access To This Company Code": Exit Sub
    Case Is = UCase("wBYy")
       Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'YY' And " + _
                                       "PMCd = '" + pv_NewValue + "'"))
        ErrMsg = "Invalid Bag Year": Exit Sub
    Case Is = UCase("wBChr")
      Cancel = (pv_NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where PTyp= 'BCHR' And " + _
                                       "PMCd = '" + pv_NewValue + "'"))
        ErrMsg = "Invalid Bag Character": Exit Sub
    Case Is = UCase("wBNo")
      'vm.2 if copy from fg bag then checks valid bag no, if from master record checks valid stkno
      If OptFgbag.Value = True Then
        '****** Sachin 3.02.0 - [03/12/07] - Bags From Current Partition only can be copied
        Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("Select BNo From Bag Where " + _
                                      "BCoCd='" + ADC("wCoCd") + "' and " + _
                                      "BYy= '" + ADC("wBYy") + "' And " + _
                                      "BChr = '" + ADC("wBChr") + "' And " + _
                                      "BNo= " + CStr(pv_NewValue) + " and BLoc= 'PFG' and BOdDmCd= '" + ADC("DmCd") + "' And BPrtKey='" + ctCurrPrtn + "'"))
        If Cancel = True Then ErrMsg = "Invalid Bag Number Or Bag Is Not In Fg For This Design": Exit Sub
      '*** Jay 2.14 *** [FG Bag Rm Copy]
      ElseIf OptMstRec.Value = True Then
        Cancel = (pv_NewValue <> 0) And (Not moCn.RecSeek("select OdDmStkNo from OrdDsg where OdCoCd='" + gs_CoCd + _
                        "' and OdDmCd= '" + ADC("DmCd") + "' and OdSfx= '' and OdDmSz= '" + ADC("DmSz") + _
                        "' and OdDmStkCoCd='" + ADC("wCoCd") + "' and OdDmStkYy='" + ADC("wBYy") + _
                        "' and OdDmStkChr = '" + ADC("wBChr") + "' and OdDmStkNo= " + CStr(pv_NewValue)))
        If Cancel = True Then ErrMsg = "Invalid Stock Number Or Stock Number Not Present For This Design Code-Size Combination": Exit Sub
      End If
    
    '*** Jay 3.1.0_1 [Dsg SetCd From Param]
    Case Is = UCase("DmSetCd")
      If moCn.RecSeek("Select 'x' from Head where HCocd = '" + ctSelfCoCd + "' and HValidDsgSetCd = 'Y' ") Then
        Cancel = Not moCn.RecSeek("Select PMCd  From Param Where PTyp= 'SETCD' And PMCd = '" + pv_NewValue + "'")
        ErrMsg = "Set Code Not Defined": Exit Sub
      End If
    '*** Jay 3.1.0_1 [Dsg SetCd From Param]
    'vm.2 validating Rm group and Rm category
    Case Is = UCase("wCpyFgBagGrp")
      Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp='ORDCPYGR' and PMCd='" + pv_NewValue + "'")
      ErrMsg = "Invalid Group By Value": Exit Sub
    Case Is = UCase("wCpyFgBagCtg")
      Cancel = pv_NewValue <> "" And Not moCn.RecSeek("select PMCd from Param " + _
                   "where PTyp= 'GLDAS' and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Raw Material Category Value": Exit Sub
  End Select
      
End Sub

Private Sub ChkFCSTW_GotFocus()
  DispMsg ChkFCSTW.ToolTipText, etInfo
End Sub

Private Sub CmdCpyFgBag_Click()
  '****** Sachin 3.02.0 - Purge Partition - Locking Previous Partition Records
  If UCase(ADC("DMPRTKEY") <> ctCurrPrtn) Then DispMsg "Cannot Copy As Design In Previous Partition", etError: Exit Sub
  Call DispFra(CpyFgBag)
  If ADC("wCoCd") = "" Then ADC("wCoCd") = gs_CoCd
  If ADC("wBYy") = "" Then ADC("wBYy") = moCn.GetFldVal("Select hYy from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")
  If ADC("wBChr") = "" Then ADC("wBChr") = moCn.GetFldVal("Select hBChr from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'")

End Sub

Private Sub CmdCpyFgBagFill_Click()
'vm.2 Procedure to show the Entire Bag Rm in grid
  Dim IdName As String, Cancel As Boolean, ErrMsg As String
  If ADC("wCoCd") = "" Or ADC("wBYy") = "" Or ADC("wBChr") = "" Or ADC("wBNo") = 0 Then
    DispMsg ("Cannot Fill Grid As Bag Details Not Proper"), etError: Exit Sub
  End If
  ms_hAddJROYN = moCn.GetFldVal("Select hAddJROYN from Head where hCoCd='" + ADC("wCoCd") + "'")
  
  'checking valid Rm group and Rm ctg options
  Select Case IdName
  Case Is = UCase("wCpyFgBagGrp")
    Cancel = Not moCn.RecSeek("select PMCd from Param where PTyp='ORDCPYGR' and PMCd= " + ADC("wCpyFgBagGrp"))
    ErrMsg = "Invalid Group By Value": Exit Sub
  Case Is = UCase("wCpyFgBagCtg")
    Cancel = ADC("wCpyFgBagCtg") <> "" And Not moCn.RecSeek("select PMCd from Param " + _
             "where PTyp= 'GLDAS' and PMCd = " + ADC("wCpyFgBagCtg"))
    ErrMsg = "Invalid Raw Material Category Value": Exit Sub
  End Select
  
  moCn.Execute ("Delete From DsgTgt " + _
                " Where DtTcTyp ='" + ADC("DmTcTyp") + "'" + _
                " And DtCd = '" + ADC("DMCd") + "'" + _
                " And DtSz ='" + ADC("DMSz") + "' ")
                          
  Dim ws_RmCtg As String, ws_GrpBy As String, ws_SqlStr As String, ws_CommStr As String, ws_CommJoinWhr As String
  Dim i, ws_RmCtgLen As Integer
  'Getting Rm ctg values as per user selection
  ws_RmCtg = ""
  If ADC("wCpyFgBagCtg") = "" Then
    ws_RmCtg = "G','P','S','L','M','X','D','C"
  Else
    ws_RmCtgLen = Len(ADC("wCpyFgBagCtg"))
    For i = 1 To ws_RmCtgLen
      ws_RmCtg = ws_RmCtg + Mid(ADC("wCpyFgBagCtg"), i, 1)
      If i <> ws_RmCtgLen Then
        ws_RmCtg = ws_RmCtg + "','"
      End If
    Next
  End If
    
  If OptFgbag.Value = True Then
    Select Case ADC("wCpyFgBagGrp")
    Case Is = "-"
      ws_GrpBy = "FrRmCtg,FrRmSCtg"
      ws_SqlStr = "Select FrRmCtg, FrRmSCtg, '*' as RmCol, '*' as qRngSz, "
    Case Is = "RMCOL"
      ws_GrpBy = "FrRmCtg, FrRmSCtg, RmCol"
      ws_SqlStr = "Select FrRmCtg, FrRmSCtg, RmCol, '*' as qRngSz, "
    Case Is = "RMCOL-SZ"
      ws_GrpBy = "FrRmCtg, FrRmSCtg, RmCol,RrRngSz"
      ws_SqlStr = "Select FrRmCtg, FrRmSCtg, RmCol, Isnull(RrRngSz,'') as qRngSz,"
    Case Is = "RMRNGSZ"
      ws_GrpBy = "FrRmCtg, FrRmSCtg, RrRngSz"
      ws_SqlStr = "Select FrRmCtg, FrRmSCtg, '*' as RmCol, Isnull(RrRngSz,'') as qRngSz,"
    Case Else
      DispMsg "No Grouping Fields Defined For Selected GroupBy Value", etError
      Exit Sub
    End Select
    
    ws_SqlStr = ws_SqlStr + "IsNull(sum(FrRmQty),0) as Qty " + _
                "from FgRm " + _
                "Join Fgd on FrFdIdNo=FdIdNo And FrCoCd= FdCoCd and FrTc= FdTc and FrYy= Fdyy and FrChr= FdChr and FrNo = FdNo And FrSr = FdSr And FrPrtKey=FdPrtKey " + _
                "Join RmMst on RmCtg=FrRmCtg and RmSCtg=FrRmSCtg And RmCd =FrRmCd " + _
                "Left outer Join RmRt on RrTcTyp ='RSZ' and RrCtg=FrRmCtg and RrSCtg=FrRmSCtg and (FrRmSz >=RrFrLn and FrRmSz <=RrToLn) " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "FrRmCtg in ('" + ws_RmCtg + "') and (Select PValue from Param where PTyp= 'TC' and PMCd= FdTc)= 'FB' and " + _
                "FdCoCd='" + ADC("wCoCd") + "' and FdBYy='" + ADC("wBYy") + "' and FdBChr ='" + ADC("wBChr") + _
                "' and FdBNo=" + CStr(ADC("wBNo")) + " Group By " + ws_GrpBy
  ElseIf OptMstRec.Value = True Then
    Select Case ADC("wCpyFgBagGrp")
    Case Is = "-"
      ws_GrpBy = "OrRmCtg,OrRmSCtg"
      ws_SqlStr = "Select OrRmCtg as FrRmCtg, OrRmSCtg as FrRmSCtg, '*' as RmCol, '*' as qRngSz, "
      ws_CommStr = "Select OrRmCtg, OrRmSCtg, '*' as RmCol, '*' as  RrRngSz, IsNull(sum(OrQty),0) as OrQty from OrdRm "
    Case Is = "RMCOL"
      ws_GrpBy = "OrRmCtg, OrRmSCtg, RmCol"
      ws_SqlStr = "Select OrRmCtg as FrRmCtg, OrRmSCtg as FrRmSCtg, RmCol, '*' as qRngSz, "
      ws_CommStr = "Select OrRmCtg, OrRmSCtg, RmCol, '*' as RrRngSz, IsNull(sum(OrQty),0) as OrQty from OrdRm "
    Case Is = "RMCOL-SZ"
      ws_GrpBy = "OrRmCtg, OrRmSCtg, RmCol,RrRngSz"
      ws_SqlStr = "Select OrRmCtg as FrRmCtg, OrRmSCtg as FrRmSCtg, RmCol, Isnull(RrRngSz,'') as qRngSz,"
      ws_CommStr = "Select OrRmCtg, OrRmSCtg, RmCol, Isnull(RrRngSz,'') as RrRngSz, IsNull(sum(OrQty),0) as OrQty from OrdRm "
    Case Is = "RMRNGSZ"
      ws_GrpBy = "OrRmCtg, OrRmSCtg, RrRngSz"
      ws_SqlStr = "Select OrRmCtg as FrRmCtg, OrRmSCtg as FrRmSCtg, '*' as RmCol, Isnull(RrRngSz,'') as qRngSz,"
      ws_CommStr = "Select OrRmCtg, OrRmSCtg, '*' as RmCol, Isnull(RrRngSz,'') as RrRngSz, IsNull(sum(OrQty),0) as OrQty from OrdRm "
    Case Else
      DispMsg "No Grouping Fields Defined For Selected GroupBy Value", etError
      Exit Sub
    End Select
    
    ws_CommJoinWhr = "Join RmMst on RmCtg =OrRmCtg and RmSCtg =OrRmSCtg and RmCd=OrRmCd " + _
                "Left outer Join RmRt on RrTcTyp ='RSZ' and RrCtg=OrRmCtg and RrSCtg=OrRmSCtg and (OrLn1 >=RrFrLn and OrLn1 <=RrToLn) " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + _
                "OrRmCtg in ('" + ws_RmCtg + "') and (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= 'JT' and " + _
                "OdDmStkCoCd='" + ADC("wCoCd") + "' and OdDmStkYy='" + ADC("wBYy") + "' and OdDmStkChr ='" + ADC("wBChr") + _
                "' and OdDmStkNo=" + CStr(ADC("wBNo")) + " Group By " + ws_GrpBy
    
    ws_SqlStr = ws_SqlStr + " IsNull(sum(OrQty),0) as Qty from " + _
                "(" + _
                  ws_CommStr + _
                  "join (Select isnull((Select OdIdNo from  vwDmStkNumMaster Where MasterYN='Y' and OdDmCd='" + ADC("DmCd") + _
                  "' And OdDmStkCoCd='" + ADC("wCoCd") + "' and OdDmStkYy='" + ADC("wBYy") + "' and " + _
                  "OdDmStkChr='" + ADC("wBChr") + "' and OdDmStkNo=" + CStr(ADC("wBNo")) + "" + _
                  "), 0) as qMstRecOdIdNo) qMstRec On qMstRecOdIdNo=OrOdIdNo " + _
                  "join OrdDsg on OdIdNo =qMstRecOdIdNo " + ws_CommJoinWhr + _
                  " Union All " + _
                  ws_CommStr + _
                  "left outer Join(Select OdIdNo from vwDmStkNumMaster where 'Y'='" + ms_hAddJROYN + "'  and MasterYN='Y' and " + _
                  "OdDmCd='" + ADC("DmCd") + "' And OdDmStkCoCd='" + ADC("wCoCd") + "' " + _
                  "and OdDmStkYy='" + ADC("wBYy") + "' and OdDmStkChr='" + ADC("wBChr") + "' " + _
                  "and OdDmStkNo=" + CStr(ADC("wBNo")) + ") qJroRec on qJroRec.OdIdNo =OrOdMstIdNo " + _
                  "join OrdDsg on OrdDsg.OdIdNo =qJroRec.OdIdNo " + ws_CommJoinWhr + _
                " ) qOrdRm Group By " + ws_GrpBy
   
  End If
  
  
  Dim wo_rsFgRmSumm As MwfLib.MDORowSet, wi_Row As Integer
  GrdCpyFgBag.Rows = 1
  Set wo_rsFgRmSumm = moCn.OpenRes(ws_SqlStr)

  If wo_rsFgRmSumm.RecCount > 0 Then
    GrdCpyFgBag.AllowAdd = True
    With GrdCpyFgBag
      Do While Not (wo_rsFgRmSumm.EOF Or wo_rsFgRmSumm.BOF)
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "DtRmCtg") = wo_rsFgRmSumm!FrRmCtg
        .Value(wi_Row, "DtRmSCtg") = wo_rsFgRmSumm!FrRmSCtg
        .Value(wi_Row, "DtCol") = wo_rsFgRmSumm!RmCol
        .Value(wi_Row, "DtRmRSz") = wo_rsFgRmSumm!qRngSz
        wo_rsFgRmSumm.MoveNext
      Loop
    End With
    GrdCpyFgBag.AllowAdd = False
  End If
  
  With GrdCpyFgBag
    If ADC("wCpyFgBagGrp") = "-" Then
      .ColProp("DtCol").Style = fgcinvisible
      .ColProp("DtRmRSz").Style = fgcinvisible
    ElseIf ADC("wCpyFgBagGrp") = "RMCOL" Then
      .ColProp("DtCol").Style = fgcNormal
      .ColProp("DtRmRSz").Style = fgcinvisible
    ElseIf ADC("wCpyFgBagGrp") = "RMCOL-SZ" Then
      .ColProp("DtCol").Style = fgcNormal
      .ColProp("DtRmRSz").Style = fgcNormal
    ElseIf ADC("wCpyFgBagGrp") = "RMRNGSZ" Then
      .ColProp("DtCol").Style = fgcinvisible
      .ColProp("DtRmRSz").Style = fgcNormal
    End If
  End With
  
End Sub

Private Sub CmdCpyFgBagGo_Click()
  If ADC("wCoCd") = "" Or ADC("wBYy") = "" Or ADC("wBChr") = "" Or ADC("wBNo") = 0 Then
    DispMsg ("Cannot Copy As Bag Details Not Proper"), etError: Exit Sub
  End If
  'vm.2 checking Rm group option and grid view
  If ADC("wCpyFgBagGrp") = "-" And (GrdCpyFgBag.ColProp("DtCol").Style = fgcNormal Or GrdCpyFgBag.ColProp("DtRmRSz").Style = fgcNormal) Then
    DispMsg "Mismatch in Grid View and Group by selection ", etError: Exit Sub
  End If
  If ADC("wCpyFgBagGrp") = "RMCOL" And (GrdCpyFgBag.ColProp("DtCol").Style = fgcinvisible Or GrdCpyFgBag.ColProp("DtRmRSz").Style = fgcNormal) Then
    DispMsg "Mismatch in Grid View and Group by selection ", etError: Exit Sub
  End If
  If ADC("wCpyFgBagGrp") = "RMCOL-SZ" And (GrdCpyFgBag.ColProp("DtCol").Style = fgcinvisible Or GrdCpyFgBag.ColProp("DtRmRSz").Style = fgcinvisible) Then
    DispMsg "Mismatch in Grid View and Group by selection ", etError: Exit Sub
  End If
  If ADC("wCpyFgBagGrp") = "RMRNGSZ" And (GrdCpyFgBag.ColProp("DtCol").Style = fgcNormal Or GrdCpyFgBag.ColProp("DtRmRSz").Style = fgcinvisible) Then
    DispMsg "Mismatch in Grid View and Group by selection ", etError: Exit Sub
  End If
  
  Call MakeRmFromFg
  Call CmdCpyFgBag_Click
End Sub

Private Sub CmdDsgMat_Click()
Call DispFra(DsgMat)
If FraDsgMatrix.Visible = True Then
    Dim wRsDmMat As MDORowSet, wPicWidth As Single, wPicHt As Single, i As Integer, wHdImg As Boolean, wClrImg As Integer
    wPicWidth = 2535: wPicHt = 2175
    ImgMatZoom.Visible = False
    For wClrImg = 1 To 9
        Call ShowPic(wPicHt, wPicWidth, "", ImgArr(wClrImg))
        LblPic(wClrImg).Caption = ""
        LblPic(wClrImg).Tag = ""
    Next
    Set wRsDmMat = moCn.OpenRes("Select PMCd, PDesc From Param Where PTyp='DMMATX'")
    If Not (wRsDmMat.BOF Or wRsDmMat.EOF) Then
        For i = 1 To wRsDmMat.RecCount
            w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , wRsDmMat!PDesc, True)
            If Dir(w3DPth) = "" Then
                w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , wRsDmMat!PDesc)
                wHdImg = False
                If Dir(w3DPth) <> "" Then
                   Call ShowPic(wPicHt, wPicWidth, w3DPth, ImgArr(i))
                   LblPic(i).Caption = wRsDmMat!PDesc + IIF(wHdImg, " [HD]", "")
                   LblPic(i).Tag = wRsDmMat!PDesc
                End If
            Else
                wHdImg = True
                Call ShowPic(wPicHt, wPicWidth, w3DPth, ImgArr(i))
                LblPic(i).Caption = wRsDmMat!PDesc + IIF(wHdImg, " [HD]", "")
                LblPic(i).Tag = wRsDmMat!PDesc
            End If
            wRsDmMat.MoveNext
        Next
    End If
End If
End Sub

'***** Bhavna 5.0.2 - 14/07/2015 new button for market card
Private Sub CmdMktPdf_Click()
  Dim wPath As String, wMktCard As String
  On Error GoTo ErrH
  wMktCard = "MktCard"
  wPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wMktCard + "\" + ADC("DmTcTyp") + " " + ADC("DmCd") + ".pdf"
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


'***** Manali 3.8.0
Private Sub CmdOpnPdf_Click()
  Dim wPath As String, wJobCard As String
  On Error GoTo ErrH
  wJobCard = "JobCard"
  wPath = moCn.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\" + wJobCard + "\" + ADC("DmTcTyp") + " " + ADC("DmCd") + ".pdf"
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
'6.1b-46
Private Sub CmdPDMLst_Click()
  Call DispFra(PDMLst)
  If FraPDMList.Visible = True Then Call ShowPDMLst

End Sub

'***** Manali 3.8.0
Private Sub CmdPushCurr_Click()
'****** Sachin 3.02.0 [26-11-2007] - Purge Partition
    If ADC("DmPrtKey") <> ctCurrPrtn Then Call PushToCurrent
End Sub

Private Sub GrdCpyFgBag_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdCpyFgBag
    Select Case ColName
    Case Is = UCase("DtRmCtg")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'RMCTG' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material Category": Exit Sub
    Case Is = UCase("DtRmSCtg")
      Cancel = (Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'RMSCTG' And " + _
                          "PMCd='" + .Value(RowNum, "DtRmCtg") + "' " + _
                           "And PSCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
    Case Is = UCase("DtCol")
      Cancel = (Not moCn.RecSeek("Select RmCol From RmMst Where RmZ='N' And " + _
                              "RmCtg ='" + .Value(RowNum, "DtRmCtg") + "' and RmSCtg='" + .Value(RowNum, "DtRmSCtg") + _
                              "' And RmPrtKey='" + ctCurrPrtn + "' and RmCol='" + NewValue + "'"))
      ErrMsg = "Invalid Raw Material Colour": Exit Sub
    Case Is = UCase("DtRmRSZ")
      Cancel = (Not moCn.RecSeek("Select RrRngSz From RmRt Where RrCtg ='" + .Value(RowNum, "DtRmCtg") + _
                              "' and RrSCtg='" + .Value(RowNum, "DtRmSCtg") + "' And RrPrtKey='" + ctCurrPrtn + _
                              "' and RrRngSz='" + NewValue + "'"))
      ErrMsg = "Invalid Range Size": Exit Sub
       
    Case Is = UCase("DtRmCd")
      If .Value(RowNum, "DtRmCtg") = "G" Or .Value(RowNum, "DtRmCtg") = "P" Or .Value(RowNum, "DtRmCtg") = "S" Or .Value(RowNum, "DtRmCtg") = "L" Or .Value(RowNum, "DtRmCtg") = "X" Then
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                                                      "RmCtg In ('G', 'P', 'S', 'L','M','X') And " + _
                                                      "RmCd = '" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'"))
        ErrMsg = "Invalid Target Raw Material Code/ Rm Code Allowed Only Of RmCtg ['G', 'P', 'S', 'L','M','X'] ": Exit Sub
      Else
        If .Value(RowNum, "DtRmCtg") = "D" Or .Value(RowNum, "DtRmCtg") = "C" Then
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                                                      "RmCtg In ('D', 'C') And " + _
                                                      "RmSCtg='" + .Value(RowNum, "DtRmSCtg") + "' And " + _
                                                      "RmCd = '" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'"))
          ErrMsg = "Invalid Target Raw Material Code/ Rm Code Allowed Only Of RmCtg ['D', 'C'] ": Exit Sub
        Else
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                             "RmCtg= '" + .Value(RowNum, "DtRmCtg") + "' And " + _
                             "RmSCtg='" + .Value(RowNum, "DtRmSCtg") + "' " + _
                             "And RmCd = '" + NewValue + "' and RmPrtKey='" + ctCurrPrtn + "'"))
          ErrMsg = "Invalid Target Raw Material Code": Exit Sub
        End If
      End If
      
    End Select
  End With
End Sub

Private Sub GrdCpyFgBag_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  Select Case UCase(ColName)
  Case Is = UCase("DtRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("DtRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdCpyFgBag.Value(RowNum, "DtRmCtg"))
  Case Is = UCase("DtCol")
    Call HlpList.PSCd("RMCOL", GrdCpyFgBag.Value(RowNum, "DtRmCtg"))
  Case Is = UCase("DtRmRSZ")
    Call HlpList.RangeSz(GrdCpyFgBag.Value(RowNum, "DtRmCtg"), GrdCpyFgBag.Value(RowNum, "DtRmSCtg"))

  Case Is = UCase("DtRmCd")
    If GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "G" Or GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "P" Or _
      GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "S" Or GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "L" Or GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "X" Then
        Call HlpList.RmCd("N", "'G','P','S','L','X'")
    Else
      If GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "D" Or GrdCpyFgBag.Value(RowNum, "DtRmCtg") = "C" Then
        Call HlpList.RmCd("N", "'D','C'", GrdCpyFgBag.Value(RowNum, "DtRmSCtg"))
      Else
        Call HlpList.RmCd("N", "'" + GrdCpyFgBag.Value(RowNum, "DtRmCtg") + "'", GrdCpyFgBag.Value(RowNum, "DtRmSCtg"))
      End If
    End If
  End Select

End Sub

Private Sub GrdCpyFgBag_InitKey(ByVal RowNum As Integer)
 '*** Initialise Key Fields
  With GrdCpyFgBag
      .Store "DtTcTyp", ADC("DmTcTyp")
      .Store "DtCd", ADC("DmCd")
      .Store "DtSz", ADC("DmSz")
  End With
End Sub

Private Sub GrdCpyFgBag_RowWhen(ByVal RowNum As Integer)
  With GrdCpyFgBag
    If .IsNew(RowNum) Then
     .Value(RowNum, "DtSr") = .MaxVal("DtSr") + 1
    End If
  End With
End Sub

Private Sub GrdCpyFgBag_SetRecSource()
  GrdCpyFgBag.RecSource = "Select * From DsgTgt" + _
                          " Where DtTcTyp ='" + ADC("DmTcTyp") + "'" + _
                          " And DtCd = '" + ADC("DMCd") + "'" + _
                          " And DtSz ='" + ADC("DMSz") + "' "
End Sub

Private Sub GrdCpyFgBag_Validate(Cancel As Boolean)
  Cancel = GrdCpyFgBag.Validate
End Sub

Private Sub GrdDsgAllow_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
'578-12 validating field values
  Select Case UCase(ColName)
  Case Is = UCase("DmwTyp")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'DMALLOW' and PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Allow Type ": Exit Sub
  Case Is = UCase("DmwCmCd")
      Cancel = Not moCn.RecSeek("Select 'x' From CustMst where " + _
                            "CmCtg= 'C' and CmCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Customer": Exit Sub
  Case Is = UCase("DmwCmGrp")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'CMGRP' and PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Group": Exit Sub
  Case Is = UCase("DmwReg")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'REGION' and PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Region": Exit Sub
  Case Is = UCase("DmwYN")
      Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                                " PTyp='YN' and PMCd= '" + NewValue + "'")
      ErrMsg = "Invalid Y/N Option": Exit Sub

  End Select
End Sub

'578-12 calling help
Private Sub GrdDsgAllow_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  If UCase(ADC("DmPrtKey")) <> ctCurrPrtn Then
    Cancel = True
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
 
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
  End If

  Select Case UCase(ColName)
  Case Is = UCase("DmwTyp")
    Call HlpList.PMCd("DMALLOW")
  Case Is = UCase("DmwCmCd")
    If UCase(GrdDsgAllow.Value(RowNum, "DmwTyp")) <> UCase("CUST") Then Cancel = True: ErrMsg = "Cannot Enter Value If Type <> 'CUST' ": Exit Sub
    Call HlpList.CustCd("C")
  Case Is = UCase("DmwCmGrp")
    If UCase(GrdDsgAllow.Value(RowNum, "DmwTyp")) <> UCase("GROUP") Then Cancel = True: ErrMsg = "Cannot Enter Value If Type <> 'GROUP' ": Exit Sub
    Call HlpList.PMCd("CMGRP")
  Case Is = UCase("DmwReg")
    If UCase(GrdDsgAllow.Value(RowNum, "DmwTyp")) <> UCase("REG") Then Cancel = True: ErrMsg = "Cannot Enter Value If Type <> 'REG' ": Exit Sub
    Call HlpList.PMCd("REGION")
  Case Is = UCase("DmwYN")
    Call HlpList.PMCd("YN")
  End Select

End Sub

'578-12
Private Sub GrdDsgAllow_GotFocus()
 If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey")) <> ctCurrPrtn) Or mf_DtLocked Then
        GrdDsgAllow.AllowAdd = False
        GrdDsgAllow.AllowDelete = False
    Else
        GrdDsgAllow.AllowAdd = True
        GrdDsgAllow.AllowDelete = True
    End If
End Sub

Private Sub GrdDsgAllow_InitKey(ByVal RowNum As Integer)
'578-12
 '*** Initialise Key Fields
  With GrdDsgAllow
      .Store "DmwDmTyp", ADC("DmTcTyp")
      .Store "DmwDmCd", ADC("DmCd")
      .Store "DmwDmSz", ADC("DmSz")
      If .IsNew(RowNum) Then
        .Store "DmwDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                            " Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' And DmPrtKey='" + ctCurrPrtn + "' ")
      End If
      
  End With
End Sub

Private Sub GrdDsgAllow_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
'578-12 if cust type selected then value can be enter only in customer, same to group and region
With GrdDsgAllow
  Select Case UCase(ColName)
  Case Is = UCase("DmwCmCd")
    If .Mode = fgmnorm Then .Value(RowNum, "DmwCmCd") = ""
  Case Is = UCase("DmwCmGrp")
    If .Mode = fgmnorm Then .Value(RowNum, "DmwCmGrp") = ""
  Case Is = UCase("DmwReg")
    If .Mode = fgmnorm Then .Value(RowNum, "DmwReg") = ""
  Case Is = UCase("DmwYN")
    If .Mode = fgmnorm Then .Value(RowNum, "DmwYN") = ""
  End Select
End With
End Sub

Private Sub GrdDsgAllow_SetRecSource()
'578-12
 GrdDsgAllow.RecSource = " Select * from DmAllow " + _
                        " Where DmwDmTyp ='" + ADC("DmTcTyp") + "'" + _
                        " And DmwDmCd = '" + ADC("DMCd") + "'" + _
                        " And DmwDmSz ='" + ADC("DMSz") + "' Order By DmwTyp"
End Sub
'578-12
Private Sub GrdDsgAllow_Validate(Cancel As Boolean)
  Cancel = GrdDsgAllow.Validate
End Sub

Private Sub GrdDsgAna_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
 'mb_Upd = True  '6.1b-46 mb_Upd check removed, it will check when Rm,lab record change
End Sub
Private Sub GrdDsgCT_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
       ' adc("DmParts") = adc("DmParts") - GrdDsgCT.Value(RowNum, "DctParts")
End Sub
Private Sub GrdDsgCT_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
With GrdDsgCT
  Select Case UCase(ColName)
    Case Is = UCase("DctRmWt")
      If .Mode = fgmnorm Then
        If UCase(SourceColName) = UCase("DctRmCd") Then
          Dim wOldRmCtg As String, wNewRmCtg As String, NewRmCtg As String, wOldSpGrav As Single, wNewSpGrav As Single
          '****** Sachin 3.02.0 - Purge Partition - Partition Key added below
          wOldRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + OldVal + "' And RmPrtKey='" + ctCurrPrtn + "' ")
          wNewRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + .Value(RowNum, "DctRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "' ")
          If Trim(wOldRmCtg) = wNewRmCtg And OldVal <> "" Then
             wOldSpGrav = moCn.GetFldVal("Select RmSpGrav from RmMst where RmCd= '" + OldVal + "' And RmPrtKey='" + ctCurrPrtn + "' ")
             wNewSpGrav = moCn.GetFldVal("Select RmSpGrav from RmMst where RmCd= '" + .Value(RowNum, "DctRmCd") + "' And RmPrtKey='" + ctCurrPrtn + "' ")
            .Value(RowNum, "DctRmWt") = MWLib.Div(.Value(RowNum, "DctRmWt") * wNewSpGrav, wOldSpGrav)
          End If
        End If
      End If
  End Select
End With

End Sub

Private Sub GrdDsgCT_RowWhen(ByVal RowNum As Integer)
  With GrdDsgCT
    If .IsNew(RowNum) Then .Value(RowNum, "DctParts") = 1
    If .Value(RowNum, "DctSr") = 0 Then _
       .Value(RowNum, "DctSr") = .MaxVal("DctSr") + 1
    If .IsNew(RowNum) Then .Value(RowNum, "DctPrtKey") = ctCurrPrtn  '****** Sachin 3.02 04-12-07 - Default value for Partition Key in RM Grid
  End With
End Sub

'Geeta***2.10
Private Sub CmdCpyRm_Click()

'****** Sachin 3.02.0 - Purge Partition - Locking Previous Partition Records
If UCase(ADC("DMPRTKEY") <> ctCurrPrtn) Then DispMsg "Cannot Copy As Design In Previous Partition", etError: Exit Sub

'CmdCpyRm Btn is Placed on NonKey Frame
'If Not moCn.RecSeek("Select DctDmCd from DsgCt where " + _
                "DctTcTyp='" + ADC("DmTcTyp") + "' and DctDmCd= '" + ADC("DmCd") + "' and DctDmSz= '" + ADC("DmSz") + "' and " + _
                "DctMainPrtYN= 'Y' And DctPrtKey='" + ctCurrPrtn + "'") Then _
'DispMsg "Main Part Does Not Exist ", etInfo: Exit Sub
Call DispFra(CpyRm)

End Sub

Private Sub CmdPrtRmCpyGo_Click()
'*** Call the Copy Routine and Close the Copy Frame
 Call CpyPrtRmDets
 Call CmdCpyRm_Click
End Sub

Private Sub GrdDsgLab_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
mb_Upd = True '4.1.2
End Sub

Private Sub GrdDsgModAvl_GotFocus()
'****** Sachin 2.13.0 - 02-06-2006 ******
'****** Sachin 3.02.0 - Purge Partition - Locking Previous Partition Records
If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey") <> ctCurrPrtn)) Or mf_DtLocked Then
        GrdDsgModAvl.AllowAdd = False
        GrdDsgModAvl.AllowDelete = False
    Else
        GrdDsgModAvl.AllowAdd = True
        GrdDsgModAvl.AllowDelete = True
    End If
'****** Sachin 2.13.0 - 02-06-2006 ******
End Sub

Private Sub GrdDsgModAvl_RowWhen(ByVal RowNum As Integer)
  '****** Sachin 2.13.0 - 02-06-2006 ******
  '*** If GrdDsgModAvl.IsNew Then DmaSr = Max + 1
  With GrdDsgModAvl
    If .Value(RowNum, "DmaSr") = 0 Then
       .Value(RowNum, "DmaSr") = .MaxVal("DmaSr") + 1
    End If
  End With
  '****** Sachin 2.13.0 - 02-06-2006 ******
End Sub

Private Sub GrdDsgModAvl_SetRecSource()
  '****** Sachin 2.13.0 - 02-06-2006 *****
  '*** Set the RecSource for GrdDsgModAvl
  GrdDsgModAvl.RecSource = " Select * from DsgModAvl " + _
                        " Where DmaTcTyp ='" + ADC("DmTcTyp") + "'" + _
                        " And DmaCd = '" + ADC("DMCd") + "'" + _
                        " And DmaSz ='" + ADC("DMSz") + "' "
  '****** Sachin 2.13.0 - 02-06-2006 *****
End Sub

'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
'If GrdDsgCT.Value(RowNum, ColName) <> NewValue Then mb_Upd = True  '6.1b-46 mb_Upd check removed, it will check when Rm,lab record change
    
  Select Case UCase(ColName)
    Case Is = UCase("DctCd"), UCase("DctSz")
      If wb_CTMod = True And GrdDsgCT.Mode = fgmnorm Then
        Dim wRsDsgCT As MwfLib.MDORowSet
        Dim ws_Cnd As String
        If UCase(ColName) = UCase("DctCd") Then
            ws_Cnd = "and DmCd= '" + NewValue + "' and DmSz= '" + GrdDsgCT.Value(RowNum, "DctSz") + "'"
        ElseIf UCase(ColName) = UCase("DctSz") Then
            ws_Cnd = "and DmCd = '" + GrdDsgCT.Value(RowNum, "DctCd") + "' and DmSz = '" + NewValue + "'"
        Else: Exit Sub
        End If
                      
        ' ****** Sachin - 3.02 (10/09/07) - Partition Key added below ******
        ' ****** Manali - 3.03 - 18/06/08 - Palladium 'L' added
        Set wRsDsgCT = moCn.OpenRes("Select DrRmCd as qDctRmCd, DmCol as qDctDmCol, DrPrdWt as qDctRmWt, " + _
                      "(Select (case when sum(case when DrRmCtg= 'D' then 1 else 0 end)> 0 then 'D' else '' end)+  " + _
                      "(case when sum(case when DrRmCtg= 'C' then 1 else 0 end)> 0 then 'C' else '' end) " + _
                      "From DsgRm " + _
                      "Where  DrDmIdNo=DmIdNo And Drtctyp = DmTcTyp and DrCd = DmCd and DrSz = DmSz and DrPrtKey=DmPrtKey and DrRmCtg in ('C','D') " + _
                      ") as qDctDc, " + _
                      "DmPrdSeq As qDctPrdSeq, '' as qDctDesc, 'N' as qDctMainPrtYn " + _
                      "From DsgMst " + _
                      "Right Outer Join DsgRm on DrDmIdNo=DmIdNo and DrTcTyp = DmTcTyp and DrCd = DmCd and DrSz = DmSz and DrPrtKey=DmPrtKey And " + _
                      "DrRmCtg in ('G','P','S','L') and DrMainMet= 'Y' " + _
                      "where DmTcTyp= 'PM' and DmPrtKey='" + ctCurrPrtn + "' " + ws_Cnd + " ")
                                                    
        With wRsDsgCT
          If Not (.EOF Or .BOF) Then
            GrdDsgCT.Value(RowNum, "DctParts") = 1
            GrdDsgCT.Value(RowNum, "DctRmCd") = !qDctRmCd
            GrdDsgCT.Value(RowNum, "DctDmCol") = !qDctDmCol
            GrdDsgCT.Value(RowNum, "DctRmWt") = !qDctRmWt
            GrdDsgCT.Value(RowNum, "DctDc") = !qDctDc
            GrdDsgCT.Value(RowNum, "DctPrdSeq") = !qDctPrdSeq
            GrdDsgCT.Value(RowNum, "DctDesc") = !qDctDesc
            GrdDsgCT.Value(RowNum, "DctMainPrtYn") = !qDctMainPrtYn
            GrdDsgCT.Value(RowNum, "DctPrtKey") = ctCurrPrtn        '****** Sachin 3.02
          Else
            GrdDsgCT.Value(RowNum, "DctParts") = 1
            GrdDsgCT.Value(RowNum, "DctRmCd") = ""
            GrdDsgCT.Value(RowNum, "DctDmCol") = ""
            GrdDsgCT.Value(RowNum, "DctRmWt") = ""
            GrdDsgCT.Value(RowNum, "DctDc") = ""
            GrdDsgCT.Value(RowNum, "DctPrdSeq") = ""
            GrdDsgCT.Value(RowNum, "DctDesc") = ""
            GrdDsgCT.Value(RowNum, "DctMainPrtYn") = ""
            GrdDsgCT.Value(RowNum, "DctPrtKey") = ctCurrPrtn        '****** Sachin 3.02
          End If
        End With
      End If
  Case Is = UCase("dctParts")
      Cancel = NewValue <= 0
      ErrMsg = "Part Quantity Can not be <= Zero": Exit Sub
  Case Is = UCase("DctRmCd")
      '****** Sachin 3.02 - Current Partition
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' " + _
                                                       "And RmCd = '" + NewValue + "' And RmPrtKey='" + ctCurrPrtn + "'"))
      ErrMsg = "Invalid Raw Material Code": Exit Sub
  Case Is = UCase("DctDmCol")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'DMCOL' and " + _
                                                      "PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Design Colour": Exit Sub
  Case Is = UCase("DctRmWt")
      Cancel = NewValue < 0
      ErrMsg = "RmWeight Can Not Be Less Then 0": Exit Sub
  Case Is = UCase("DctDc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                " PTyp='CTDC' and PMCd= '" + NewValue + "'"))
      ErrMsg = "Invalid Component DC": Exit Sub
  Case Is = UCase("DctPrdSeq")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                           "vPTyp= 'PRDSEQ' and vPMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Production Sequence": Exit Sub
  Case Is = UCase("DctMainPrtYn")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("select PMCd from Param where " + _
                                " PTyp='YN' and PMCd= '" + NewValue + "'"))
      ErrMsg = "Invalid Main Part Y/N Option": Exit Sub
  End Select
End Sub

'*** Jay 3.4[DP]
Private Sub GrdDsgPrm_GotFocus()
'*** Jay 3.4[DP]
    '****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey")) <> ctCurrPrtn) Or mf_DtLocked Then
        GrdDsgPrm.AllowAdd = False
        GrdDsgPrm.AllowDelete = False
    Else
        GrdDsgPrm.AllowAdd = True
        GrdDsgPrm.AllowDelete = True
    End If
End Sub

'*** Jay 3.4[DP]
Private Sub GrdDsgPrmDet_GotFocus()
'*** Jay 3.4[DP]
    '****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey")) <> ctCurrPrtn) Or mf_DtLocked Then
        GrdDsgPrmDet.AllowAdd = False
        GrdDsgPrmDet.AllowDelete = False
    Else
        GrdDsgPrmDet.AllowAdd = True
        GrdDsgPrmDet.AllowDelete = True
    End If
End Sub

Private Sub GrdDsgRm_OnDelete(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
mb_Upd = True '4.1.2
End Sub

'*** Jay 3.01 *** [NAZ]
Private Sub GrdDsgRm_RowSave(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
  If UCase(GrdDsgRm.Value(RowNum, "DrRmCtg")) = "D" Or UCase(GrdDsgRm.Value(RowNum, "DrRmCtg")) = "C" Then _
    Cancel = moCn.RecSeek("Select 'x' from RmRt where RrTcTyp= 'NAS' and " + _
                          "RrCtg= '" + GrdDsgRm.Value(RowNum, "DrRmCtg") + "' and " + _
                          "RrSCtg= '" + GrdDsgRm.Value(RowNum, "DrRmSCtg") + "' and " + _
                          "RrFrLn= " + CStr(GrdDsgRm.Value(RowNum, "DrLn1")) + " and " + _
                          "RrToLn= " + CStr(GrdDsgRm.Value(RowNum, "DrLn2")))
  If Cancel = True Then ErrMsg = "Cannot Enter This Rm Size Combination As It Is Not available Size For This Rm Sub Ctg"
  
  Cancel = ValidateStkSize(GrdDsgRm.Value(RowNum, "DrLn1"), GrdDsgRm.Value(RowNum, "DrLn2"), GrdDsgRm.Value(RowNum, "DrLn3"), ErrMsg)
End Sub

Private Sub GrdDsgTgt_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** DtRmCtg          valid PMCd('RMCTG')
  '*** DtRmSCtg         valid PSCd('RMSCTG', DtRmCtg) or ''
  '*** DtRmCd           valid RmCd from RmMst(RmCtg= DtRmCtg, RmSCtg= DtRmCtg, RmZ= 'N') or ''
  
  With GrdDsgTgt
    Select Case ColName
    Case Is = UCase("DtRmCtg")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'RMCTG' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material Category": Exit Sub
    Case Is = UCase("DtRmSCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'RMSCTG' And " + _
                           "PMCd='" + .Value(RowNum, "DtRmCtg") + "' " + _
                           "And PSCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
    Case Is = UCase("DtRmCd")
      '****** Sachin 3.02.0 - Purge Partition - Partition Key added below
      ' ***** Manali 3.5.0 - 06/12/08 - Change of RmCd within RmCtgs allowed for Main Metals
      If .Value(RowNum, "DtRmCtg") = "G" Or .Value(RowNum, "DtRmCtg") = "P" Or .Value(RowNum, "DtRmCtg") = "S" Or .Value(RowNum, "DtRmCtg") = "L" Then
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                             "RmCtg In ('G', 'P', 'S', 'L') And RmCd = '" + NewValue + "' And RmPrtKey='" + ADC("DmPrtKey") + "'"))
        ErrMsg = "Invalid Target Raw Material Code/ Rm Code only allowed for RmCtg ('G', 'P', 'S', 'L')": Exit Sub
      Else
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                             "RmCtg= '" + .Value(RowNum, "DtRmCtg") + "' And " + _
                             "RmSCtg='" + .Value(RowNum, "DtRmSCtg") + "' " + _
                             "And RmCd = '" + NewValue + "' And RmPrtKey='" + ADC("DmPrtKey") + "'"))
        ErrMsg = "Invalid Target Raw Material Code": Exit Sub
      End If
      
'      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
'                           "RmCtg= '" + .Value(RowNum, "DtRmCtg") + "' And " + _
'                           "RmSCtg='" + .Value(RowNum, "DtRmSCtg") + "' " + _
'                           "And RmCd = '" + NewValue + "' And RmPrtKey='" + ADC("DmPrtKey") + "'"))
'      ErrMsg = "Invalid Target Raw Material Code": Exit Sub
    End Select
  End With
End Sub

Private Sub GrdDsgTgt_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** In the help of DtRmCd do not show the Custom Codes
  '    The first param is Custom Code (Y/N), the 2nd is the RmCtg and third is RmSCtg
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit"
    '*** Jay 2.14 *** (SEO Change)
  End If
  Select Case UCase(ColName)
  Case Is = UCase("DtRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("DtRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdDsgTgt.Value(RowNum, "DtRmCtg"))
  Case Is = UCase("DtRmCd")
    ' **** Manali 3.5.0 - 06/12/08 - Change of RmCd within RmCtgs allowed for Main Metals
    If GrdDsgTgt.Value(RowNum, "DtRmCtg") = "G" Or GrdDsgTgt.Value(RowNum, "DtRmCtg") = "P" Or _
      GrdDsgTgt.Value(RowNum, "DtRmCtg") = "S" Or GrdDsgTgt.Value(RowNum, "DtRmCtg") = "L" Then
        Call HlpList.RmCd("N", "'G','P','S','L'")
    Else
        Call HlpList.RmCd("N", "'" + GrdDsgTgt.Value(RowNum, "DtRmCtg") + "'", GrdDsgTgt.Value(RowNum, "DtRmSCtg"))
    End If
  End Select
End Sub
Private Sub GrdDsgTgt_InitKey(ByVal RowNum As Integer)
 '*** Initialise Key Fields
  With GrdDsgTgt
      .Store "DtTcTyp", ADC("DmTcTyp")
      .Store "DtCd", ADC("DmCd")
      .Store "DtSz", ADC("DmSz")
  End With
End Sub

Private Sub GrdDsgTgt_RowWhen(ByVal RowNum As Integer)
    '*** If GrdDsgTgt.IsNew Then DtSr= Max + 1
  With GrdDsgTgt
    If .IsNew(RowNum) Then
     .Value(RowNum, "DtSr") = .MaxVal("DtSr") + 1
    End If
  End With
End Sub
Private Sub GrdDsgTgt_SetRecSource()
        GrdDsgTgt.RecSource = "Select * From DsgTgt" + _
                                " Where DtTcTyp ='" + ADC("DmTcTyp") + "'" + _
                                " And DtCd = '" + ADC("DMCd") + "'" + _
                                " And DtSz ='" + ADC("DMSz") + "' "
 End Sub
'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit":Exit Sub
    '*** Jay 2.14 *** (SEO Change)
  End If
  Select Case UCase(ColName)
  Case Is = UCase("DctCd")
    Call HlpList.DmCd("PM")
  Case Is = UCase("DctSz")
    Call HlpList.DmSz("PM", GrdDsgCT.Value(RowNum, "DctCd"))
  Case Is = UCase("DctChr")
    If GrdDsgCT.IsNew(RowNum) Then Call HlpList.PMCd("BCHR", "'C'")
  Case Is = UCase("DctRmCd")
  ' ***** Manali 3.03 - 18/06/08 - Palladou'L' added
    Call HlpList.RmCd("N", "'G','P','S','L'")
  Case Is = UCase("DctDmCol")
    Call HlpList.PMCd("DMCOL")
  Case Is = UCase("DctDC")
    Call HlpList.PMCd("CTDC")
  Case Is = UCase("DctPrdSeq")
      Call HlpList.vPMCd(gs_CoCd, "PRDSEQ")
  Case Is = UCase("DctMainPrtYn")
      Call HlpList.PMCd("YN")
  End Select
End Sub

'*** Jay 2.13(CT) ***
'****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
Private Sub GrdDsgCT_GotFocus()
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey")) <> ctCurrPrtn) Or mf_DtLocked Then
        GrdDsgCT.AllowAdd = False
        GrdDsgCT.AllowDelete = False
    Else
        GrdDsgCT.AllowAdd = True
        GrdDsgCT.AllowDelete = True
    End If
End Sub
'Geeta***2.10
Private Sub GrdDsgRm_GotFocus()
    '****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey")) <> ctCurrPrtn) Or mf_DtLocked Then
        GrdDsgRm.AllowAdd = False
        GrdDsgRm.AllowDelete = False
    Else
        GrdDsgRm.AllowAdd = True
        GrdDsgRm.AllowDelete = True
    End If
End Sub
Private Sub GrdDsgRm_RowWhen(ByVal RowNum As Integer)
  '*** If GrdDsgRm.IsNew Then DrSr = Max + 1
  With GrdDsgRm
    If .Value(RowNum, "DrSr") = 0 Then _
       .Value(RowNum, "DrSr") = .MaxVal("DrSr") + 1
    If .IsNew(RowNum) Then .Value(RowNum, "DrPrtKey") = ctCurrPrtn  '****** Sachin 3.02 04-12-07 - Default value for Partition Key in RM Grid
  End With
End Sub
Private Sub GrdDsgRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  If ADC("DmPrtKey") <> ctCurrPrtn Then
    Cancel = True: GrdDsgRm.AllowAdd = False: GrdDsgRm.AllowDelete = False
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit"
    '*** Jay 2.14 *** (SEO Change)
  End If
  
  '*** Set Helps For Each Field
  Select Case UCase(ColName)
  Case Is = UCase("DrRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("DrRmSCtg")
    Call HlpList.PSCd("RMSCTG", GrdDsgRm.Value(RowNum, "DrRmCtg"))
    
  Case Is = UCase("DrLn1")
    '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmLn Master]******
    If mb_VldRmLn = True And (GrdDsgRm.Value(RowNum, "DrRmCtg") = "D" Or GrdDsgRm.Value(RowNum, "DrRmCtg") = "C") Then
      Dim ws_Arr() As String, ws_DrLn1 As String
      ws_DrLn1 = GrdDsgRm.Value(RowNum, "DrLn1")
      ws_Arr = Split("VrsLn2,DrLn2", ",")
      Call HlpList.OrLn1(GrdDsgRm.Value(RowNum, "DrRmCtg"), moCn.GetFldVal("Select RmId From RmMst Where RmCtg='" + GrdDsgRm.Value(RowNum, "DrRmCtg") + "' And RmCd='" + GrdDsgRm.Value(RowNum, "DrRmCd") + "'"), ws_Arr)
    End If
    '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master]******
  Case Is = UCase("DrLn3")
    Cancel = ColWhenStkSz(3, GrdDsgRm.Value(RowNum, "DrRmCtg"), GrdDsgRm.Value(RowNum, "DrRmSCtg"), ErrMsg)
  Case Is = UCase("DrRmCd")
    Call HlpList.RmCd("N", "'" + GrdDsgRm.Value(RowNum, "DrRmCtg") + "'", GrdDsgRm.Value(RowNum, "DrRmSCtg"))
  Case Is = UCase("DrSetSCd")
  'Manoj
    If InStr(1, "CD", GrdDsgRm.Value(RowNum, "DrRmCtg")) <= 0 Then
      Cancel = True
      ErrMsg = "Can not Edit Setting Type for this RM Category"
    Else
      Call HlpList.PSCd("LABSCD", "SET")
    End If
  Case Is = UCase("DrSubShp")
    Call HlpList.PMCd("SUBSHP") '5717 subshp applicable for metals also
  Case Is = UCase("DrAlyCd")
    Call HlpList.RmCd("N", "'A'")
  Case Is = UCase("DrMainMet")
    Call HlpList.PMCd("YN")
  Case Is = UCase("DrRmPtr")
    If UCase(moCn.GetFldVal("Select PValue2 from Param where PTyp= 'RMSCTG' and " + _
            "PMCd= '" + GrdDsgRm.Value(RowNum, "DrRmCtg") + "' and " + _
            "PSCd= '" + GrdDsgRm.Value(RowNum, "DrRmSCtg") + "' ")) = "N" Then _
      Cancel = True: ErrMsg = "Cannot Edit Rm Pointer": Exit Sub
  Case Is = UCase("wShowPrdRmDets")
    Call HlpList.PMCd("YN")
  End Select
End Sub
Private Sub GrdDsgRm_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** DrRmPtr on DrRmCd,DrLn1,DrLn2          Normal
  
  '*** DrWt on DrRmCd,DrLn1,DrLn2,DrQty       Normal
      'If the SourceColName= DrRmCd (i.e. if DrRmCd has triggered the recalc) and
      'RmCtg= 'G' or 'P' then
      '   if Old RmCtg= New RmCtg then
      '       DrWt= DrWt* (SpGrav of New RmCd) / (SpGrav of Old RmCd)
      '   else
      '       DrWt = 0
      'Else
      '   if SourceColName= DrRmCd and OldVal<> '0' then
      '       calculate the Wt as DrWt= GetWtFrSz() only if the RmCtg has changed or
      '       if the Wt is 0
      '   else if any other dependent field has changed then DrWt = GetWtFrSz()
  '*** DrSubShp on DrRmCtg, DrRmSCtg          Normal
      'DrSubShp = ''
  '*** DrAlyCd on DrRmCd                      Normal
      'DrAlyCd = CaAlyCd from CustAly(CaCtg= 'C', CaCd= 'ZSELF', CaRmCd= DrRmCd, CaDmCol= DmCol)
  '*** DrWsQty on DrQty, DrRmCtg              Normal
      'If DrRmCtg='D' or 'C' then DrQty
  '*** DrHsQty on DrWsQty, DrQty              Normal
      'If DrRmCtg='D' or 'C' then DrQty- DrWsQty
      
  Dim wOldRmCtg As String, wRmSTWGrp As String, wFndPcWt As Double
  
  Dim wRmDetArr() As String
    
  With GrdDsgRm
    Select Case UCase(ColName)
    Case Is = UCase("DrRmPtr")
      If .Mode = fgmnorm Then
        'If .Value(RowNum, "DrRmCtg") = "C" Or .Value(RowNum, "DrRmCtg") = "D" And _
        '   UCase(moCn.GetFldVal("Select PValue2 from Param where PTyp= 'RMSCTG' and " + _
        '   "PMCd= '" + .Value(RowNum, "DrRmCtg") + "' and PSCd= '" + .Value(RowNum, "DrRmSCtg") + "' ")) = "Y" Then
          wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "'", ""))
        
        If (.Value(RowNum, "DrRmCtg") = "C" Or .Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "G" Or .Value(RowNum, "DrRmCtg") = "P" Or .Value(RowNum, "DrRmCtg") = "S" Or .Value(RowNum, "DrRmCtg") = "L" Or .Value(RowNum, "DrRmCtg") = "X" Or .Value(RowNum, "DrRmCtg") = "M") And _
           wRmSTWGrp <> "" And UCase(moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'")) = "Y" Then
           
          If UCase(SourceColName) = UCase("DrRmCd") And CStr(OldVal) <> "0" Then

             '***Geeta***3.02**
             Dim wRmDetOld, wOldRmSCtg, wRmSTWGrpOld  As String
             wRmDetOld = moCn.GetFldVal("Select RmCtg+','+RmSCtg+ ','+ RmSTWGrp from RmMst where RmCd= '" + OldVal + "' " + IIF(.Value(RowNum, "DrPrtKey") = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", ""))
              If wRmDetOld <> "" Then
                wRmDetArr = Split(wRmDetOld, ",")
                wOldRmCtg = wRmDetArr(0)
                wOldRmSCtg = wRmDetArr(1)
                wRmSTWGrpOld = wRmDetArr(2)
              End If

             If Trim(wOldRmCtg) <> .Value(RowNum, "DrRmCtg") Or Trim(wOldRmSCtg) <> .Value(RowNum, "DrRmSCtg") Or wRmSTWGrp <> wRmSTWGrpOld Or .Value(RowNum, "DrRmPtr") = 0 Then
             '***Geeta***3.02**
              .Value(RowNum, "DrRmPtr") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
                                          .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
                                          .Value(RowNum, "DrLn2"), wRmSTWGrp)
              End If
          Else
            .Value(RowNum, "DrRmPtr") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
                                        .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
                                        .Value(RowNum, "DrLn2"), wRmSTWGrp)
          End If
        Else
        .Value(RowNum, "DrRmPtr") = 0
        End If
      End If
    Case Is = UCase("DrPrdWt")
      If .Mode = fgmnorm Then
        'Sachin 2.12  12/12/05  - SJM - Silver Included Below
        ' ***** Manali 3.03 -18/06/08 - Palladium included
        ' ***** Manali 3.8.0 - Default Wt for GPSL From RmMst Based On Qty
        wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "'", ""))
          wFndPcWt = moCn.GetFldVal("Select RmFndWt From RmMst Where RmCd='" + .Value(RowNum, "DrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
        
        If UCase(SourceColName) = UCase("DrRmCd") And (.Value(RowNum, "DrRmCtg") = "G" Or _
           .Value(RowNum, "DrRmCtg") = "P" Or .Value(RowNum, "DrRmCtg") = "S" Or .Value(RowNum, "DrRmCtg") = "L" Or .Value(RowNum, "DrRmCtg") = "X" Or .Value(RowNum, "DrRmCtg") = "M") And wRmSTWGrp = "" Then
          Dim wOldSpGrav As Single, wNewSpGrav As Single
          wOldRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + OldVal + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
          ' ***** Manali 3.8.0 - Default Wt for GPSL From RmMst Based On Qty
          If wFndPcWt <> 0 Then
            .Value(RowNum, "DrPrdWt") = .Value(RowNum, "DrPrdQty") * wFndPcWt
          ' ***** Manali 3.8.0 - Default Wt for GPSL From RmMst Based On Qty
          ElseIf (Trim(wOldRmCtg) = "G" Or Trim(wOldRmCtg) = "P" Or Trim(wOldRmCtg) = "S" Or Trim(wOldRmCtg) = "L") And OldVal <> "" Then
'          Trim(wOldRmCtg) = .Value(RowNum, "DrRmCtg") And OldVal <> "" Then
            wOldSpGrav = moCn.GetFldVal("Select RmSpGrav from RmMst where RmCd= '" + OldVal + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
            wNewSpGrav = moCn.GetFldVal("Select RmSpGrav from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
            .Value(RowNum, "DrPrdWt") = MWLib.Div(.Value(RowNum, "DrPrdWt") * wNewSpGrav, wOldSpGrav)
          Else
            .Value(RowNum, "DrPrdWt") = 0
          End If
        Else
          'Sachin 2.12  12/12/05  - SJM - Silver Included Below
          ' ***** Manali 3.03 -18/06/08 - Palladium included
          If wRmSTWGrp <> "" And (.Value(RowNum, "DrRmCtg") = "G" Or .Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "C" Or .Value(RowNum, "DrRmCtg") = "P" Or .Value(RowNum, "DrRmCtg") = "S" Or .Value(RowNum, "DrRmCtg") = "L" Or .Value(RowNum, "DrRmCtg") = "X" Or .Value(RowNum, "DrRmCtg") = "M") Then
            If .Value(RowNum, "DrRmPtr") = 0 Then
              wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
              If UCase(SourceColName) = UCase("DrRmCd") And CStr(OldVal) <> "0" Then
                wOldRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + OldVal + "' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
                    If (Trim(wOldRmCtg) <> .Value(RowNum, "DrRmCtg") Or .Value(RowNum, "DrPrdWt") = 0) Then    '3.11.2 - 'X' Recalc Issue
                      .Value(RowNum, "DrPrdWt") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
                                             .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
                                             .Value(RowNum, "DrLn2"), wRmSTWGrp) * .Value(RowNum, "DrPrdQty")
                    Else
                        .Value(RowNum, "DrPrdWt") = IIF(wFndPcWt <> 0, .Value(RowNum, "DrPrdQty") * wFndPcWt, .Value(RowNum, "DrPrdWt"))
                    End If
               Else
                  wFndPcWt = moCn.GetFldVal("Select RmFndWt From RmMst Where RmCd='" + .Value(RowNum, "DrRmCd") + "'" + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey='" + ctCurrPrtn + "' ", ""))
                  If wRmSTWGrp = "" And (.Value(RowNum, "DrRmCtg") = "X" And wFndPcWt <> 0) Then
                    .Value(RowNum, "DrPrdWt") = .Value(RowNum, "DrPrdQty") * wFndPcWt
                  Else
                    .Value(RowNum, "DrPrdWt") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
                                         .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
                                         .Value(RowNum, "DrLn2"), wRmSTWGrp) * .Value(RowNum, "DrPrdQty")
                  End If
              End If
            Else
              .Value(RowNum, "DrPrdWt") = .Value(RowNum, "DrRmPtr") * .Value(RowNum, "DrPrdQty")
            End If
          ' ***** Manali 3.8.0 - Default Wt for GPSL From RmMst Based On Qty
          Else
            If wFndPcWt <> 0 Then .Value(RowNum, "DrPrdWt") = .Value(RowNum, "DrPrdQty") * wFndPcWt
          ' ***** Manali 3.8.0 - Default Wt for GPSL From RmMst Based On Qty
          End If
        End If
      End If

'*** Old Code (Bef 1.10 Rel)
'    Case Is = UCase("DrWt")
'      If .Mode = fgmnorm Then
'        Dim wOldRmCtg As String, wRmSTWGrp As String
'        If UCase(SourceColName) = UCase("DrRmCd") And (.Value(RowNum, "DrRmCtg") = "G" Or _
'           .Value(RowNum, "DrRmCtg") = "P") Then
'          Dim wOldPurityRt As Single, wNewPurityRt As Single
'          wOldRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + OldVal + "' ")
'          If Trim(wOldRmCtg) = .Value(RowNum, "DrRmCtg") And OldVal <> "" Then
'            wOldPurityRt = moCn.GetFldVal("Select RmPurityRt from RmMst where RmCd= '" + OldVal + "' ")
'            wNewPurityRt = moCn.GetFldVal("Select RmPurityRt from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' ")
'            .Value(RowNum, "DrWt") = MWLib.Div(.Value(RowNum, "DrWt") * wNewPurityRt, wOldPurityRt)
'          Else
'            .Value(RowNum, "DrWt") = 0
'          End If
'        Else
'          wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' ")
'          If UCase(SourceColName) = UCase("DrRmCd") And CStr(OldVal) <> "0" Then
'            wOldRmCtg = moCn.GetFldVal("Select RmCtg from RmMst where RmCd= '" + OldVal + "' ")
'            If Trim(wOldRmCtg) <> .Value(RowNum, "DrRmCtg") Or .Value(RowNum, "DrWt") = 0 Then _
'              .Value(RowNum, "DrWt") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
'                                     .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
'                                     .Value(RowNum, "DrLn2"), wRmSTWGrp) * .Value(RowNum, "DrQty")
'          Else
'            .Value(RowNum, "DrWt") = GetWtFrSz(.Value(RowNum, "DrRmCtg"), _
'                                     .Value(RowNum, "DrRmSCtg"), .Value(RowNum, "DrLn1"), _
'                                     .Value(RowNum, "DrLn2"), wRmSTWGrp) * .Value(RowNum, "DrQty")
'          End If
'        End If
'      End If
'*** Old Code (Bef 1.10 Rel)


'*** Old Code
'    Case Is = UCase("DrWt")
'      If .Mode = fgmnorm Then
'        Select Case UCase(SourceColName)
'        Case Is = UCase("DrRmCd")
'          If OldVal <> "" Then
'            Dim wOldPurityRt As Single, wNewPurityRt As Single
'            wOldPurityRt = moCn.GetFldVal("Select RmPurityRt from RmMst where RmCd= '" + OldVal + "' ")
'            wNewPurityRt = moCn.GetFldVal("Select RmPurityRt from RmMst where RmCd= '" + .Value(RowNum, "DrRmCd") + "' ")
'            .Value(RowNum, "DrWt") = MWLib.Div(.Value(RowNum, "DrWt") * wNewPurityRt, wOldPurityRt)
'          End If
'        End Select
'      End If
'*** Old Code

    Case Is = UCase("DrQty")
      If .Mode = fgmnorm Then
        .Value(RowNum, "DrQty") = .Value(RowNum, "DrPrdQty")
      End If
    Case Is = UCase("DrWt")
      If .Mode = fgmnorm Then
        .Value(RowNum, "DrWt") = .Value(RowNum, "DrPrdWt")
      End If
    Case Is = UCase("DrSubShp")
      If .Mode = fgmnorm Then .Value(RowNum, "DrSubShp") = ""
    Case Is = UCase("DrAlyCd")
      If .Mode = fgmnorm Then
        .Value(RowNum, "DrAlyCd") = moCn.GetFldVal("Select CaAlyCd from " + _
                                    "CustAly where CaCtg= 'C' and " + _
                                    "CaCd= '" + ctSelfCmCd + "' and CaRmCd= '" + .Value(RowNum, "DrRmCd") + "' " + _
                                    "and CaDmCol= '" + ADC("DmCol") + "'")
      End If
    Case Is = UCase("DrWsQty")
      Dim wWsQty As Single
      If .Mode = fgmnorm Then
        wWsQty = 0
        If .Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "C" Then
          wWsQty = .Value(RowNum, "DrQty")
        End If
        .Value(RowNum, "DrWsQty") = wWsQty
      End If
    Case Is = UCase("DrHsQty")
      Dim wHsQty  As Single
      If .Mode = fgmnorm Then
        wHsQty = 0
        If .Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "C" Then
          wHsQty = .Value(RowNum, "DrQty") - .Value(RowNum, "DrWsQty")
          If wHsQty < 0 Then wHsQty = 0
        End If
        .Value(RowNum, "DrHsQty") = wHsQty
      End If
    Case Is = UCase("DrSetSCd")
      If .Mode = fgmnorm Then
        If .Value(RowNum, "DrRmCtg") <> "D" And .Value(RowNum, "DrRmCtg") <> "C" Then
          GrdDsgRm.Value(RowNum, "DrSetSCd") = ""
        End If
      End If
    End Select
  End With
End Sub
Private Sub GrdDsgRm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** DrRmCtg          valid PMCd('RMCTG')
  '*** DrRmSCtg         valid PSCd('RMSCTG', DrRmCtg) or ''
  '*** DrRmCd           valid RmCd from RmMst(RmCtg= DrRmCtg, RmSCtg= DrRmSCtg, RmZ='N') or ''
  '*** DrSetSCd         valid PSCd('LABSCD', 'SET') or ''
  '*** DrSubShp         valid PMCd('SUBSHP')
  '*** DrAlyCd          If DrRmCtg='P' or 'G' then valid RmCd from RmMst(RmCtg= 'A', RmZ='N') or ''
                       'If DrRmCtg<>'P' and DrRmCtg<> 'G' then ''
  '*** DrMainMet        valid PMCd('YN')
                       'If DrRmCtg='P' or 'G' or 'M' only then 'Y' is allowed
  '*** DrWsQty          <= DrQty
    
  
  
  With GrdDsgRm
    If .Value(RowNum, ColName) <> NewValue Then mb_Upd = True
    
    Select Case UCase(ColName)
    Case Is = UCase("DrRmCtg")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'RMCTG' And PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
      'Manoj
      If NewValue <> "C" And NewValue <> "D" Then GrdDsgRm.Value(RowNum, "DrSetSCd") = ""
    
    Case Is = UCase("DrRmSCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                           "PTyp= 'RMSCTG' And " + _
                           "PMCd='" + .Value(RowNum, "DrRmCtg") + "' " + _
                           "And PSCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
      
    '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master] ******
    Case Is = UCase("DrLn1")
      If mb_VldRmLn = True And (.Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "C") Then
        Cancel = Not moCn.RecSeek(" Select 'x' From VldRmLn Where VrsRmCtg= '" + .Value(RowNum, "DrRmCtg") + "' And " + _
                             " VrsRmID=(Select RmId From RmMst Where RmCtg='" + .Value(RowNum, "DrRmCtg") + "' And RmCd= '" + .Value(RowNum, "DrRmCd") + "') And VrsLn1 = " + CStr(NewValue) + " ")
        ErrMsg = "Raw Material Ln Not Defined": Exit Sub
      End If
    
    Case Is = UCase("DrLn2")
      If mb_VldRmLn = True And (.Value(RowNum, "DrRmCtg") = "D" Or .Value(RowNum, "DrRmCtg") = "C") Then
        Cancel = Not moCn.RecSeek("Select 'x' From VldRmLn Where VrsRmCtg= '" + .Value(RowNum, "DrRmCtg") + "' And " + _
                             " VrsRmID=(Select RmId From RmMst Where RmCtg='" + .Value(RowNum, "DrRmCtg") + "' And RmCd='" + .Value(RowNum, "DrRmCd") + "') And VrsLn1 = " + CStr(.Value(RowNum, "DrLn1")) + " And VrsLn2 = " + CStr(NewValue))
        ErrMsg = "Raw Material Ln1-Ln2 Combination Not Defined": Exit Sub
      End If
    '****** Sachin 3.3.0 (28-06-08) - [28. Valid RmSz Master] ******
      
    Case Is = UCase("DrRmCd")
      '****** Sachin 3.02 14-12-07 - Partition Key added below
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmZ='N' And " + _
                           "RmCtg= '" + .Value(RowNum, "DrRmCtg") + "' And " + _
                           "RmSCtg='" + .Value(RowNum, "DrRmSCtg") + "' " + _
                           "And RmCd = '" + NewValue + "' And RmPrtKey='" + ctCurrPrtn + "'"))
      '*** Jay 2.13(CT) ***
      If Cancel = True Then ErrMsg = "Invalid Raw Material Code/Raw Material Code Not In Current Partition": Exit Sub
      If Cancel = False And ADC.MenuCd = "DM" And ADC("DMSZ") = "" And wb_CTMod = True Then
      ' ***** Manali 3.03 -18/06/08 - Palladium included
        If .Value(RowNum, "DrRmCtg") = "G" Or .Value(RowNum, "DrRmCtg") = "P" Or .Value(RowNum, "DrRmCtg") = "S" _
          Or .Value(RowNum, "DrRmCtg") = "L" Then
            Dim i As Integer
            i = 1
            Do While i < GrdDsgCT.Rows
              If GrdDsgCT.Value(i, "DctRmCd") = .Value(RowNum, "DrRmCd") Then GrdDsgCT.Value(i, "DctRmCd") = NewValue
              i = i + 1
            Loop
            GrdDsgRm.AddSave "GrdDsgCT", asAll
        End If
      End If
      '*** Jay 2.13(CT) ***
    Case Is = UCase("DrSetSCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                           "PTyp= 'LABSCD' And PMCd='SET' " + _
                           "And PSCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Setting Code": Exit Sub
    Case Is = UCase("DrSubShp")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'SUBSHP' And PMCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Sub Shape": Exit Sub
    Case Is = UCase("DrAlyCd")
      '****** Sachin 2.12 12/12/05  - SJM - Silver included below
      ' ***** Manali 3.03 -18/06/08 - Palladium included
      If .Value(RowNum, "DrRmCtg") = "P" Or .Value(RowNum, "DrRmCtg") = "G" Or .Value(RowNum, "DrRmCtg") = "S" _
        Or .Value(RowNum, "DrRmCtg") = "L" Then
          Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' And ", "") + " RmZ='N' And " + _
                               "RmCtg= 'A' And RmCd = '" + NewValue + "'"))
          If Cancel = True Then ErrMsg = "Invalid Alloy Code": Exit Sub
      Else
        If NewValue <> "" Then Cancel = True: ErrMsg = "Alloy Code Should Not Be Entered": Exit Sub
      End If
    '### Should DrMainMet Be 'Y' only for Recs with DrRmCtg = 'P' or 'G'  or 'M'
    '### If so the check has not been put in the trigger
    Case Is = UCase("DrMainMet")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'YN' And PMCd = '" + NewValue + "'")
      If Cancel = True Then: ErrMsg = "Invalid Main Metal Y/N Option": Exit Sub
      'Sachin 2.12 12/12/05 - SJM - Silver Included below
      ' ***** Manali 3.03 -18/06/08 - Palladium included
      If NewValue = "Y" And .Value(RowNum, "DrRmCtg") <> "G" And _
         .Value(RowNum, "DrRmCtg") <> "P" And .Value(RowNum, "DrRmCtg") <> "S" And .Value(RowNum, "DrRmCtg") <> "L" _
         And .Value(RowNum, "DrRmCtg") <> "M" Then _
         Cancel = True: ErrMsg = "This Raw Material Cannot Be Main Metal": Exit Sub
    Case Is = UCase("DrWsQty")
      If NewValue > GrdDsgRm.Value(RowNum, "DrQty") Then
        Cancel = True: ErrMsg = "Wax Set Quantity Cannot Be Greater Than RM Quantity"
      End If
    End Select
  End With
End Sub

Private Sub GrdDsgModAvl_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  
 '****** Sachin 2.13.0 - 02-06-2006 ******
 Select Case UCase(ColName)
  Case Is = UCase("DmaAvlSz")
      Cancel = NewValue <> "" And Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'DMSZ' and PMCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Design Size": Exit Sub
    
      Cancel = moCn.RecSeek("Select 'x' From DsgModAvl Where DmaTcTyp='" + ADC("DmTcTyp") + "'" + _
                                  " And DmaCd='" + ADC("DmCd") + "' And DmaSz='" + ADC("DmSz") + "' And DmaAvlSz='" + NewValue + "' ")
      If Cancel = True Then ErrMsg = "Size Already Exists in Available Sizes List for this Design": Exit Sub
 End Select
 '****** Sachin 2.13.0 - 02-06-2006 ******
End Sub

Private Sub GrdDsgModAvl_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  If UCase(ADC("DmPrtKey")) <> ctCurrPrtn Then
    Cancel = True
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  
  '****** Sachin 2.13.0 - 02-06-2006 ******
  '*** Set Helps For Each Field
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit"
    '*** Jay 2.14 *** (SEO Change)
  End If
  
  Select Case UCase(ColName)
  Case Is = UCase("DmaAvlSz")
    Call HlpList.PMCd("DMSZ")
  End Select
  '****** Sachin 2.13.0 - 02-06-2006 ******
End Sub

Private Sub GrdDsgRm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgRm
      .Store "DrTcTyp", ADC("DmTcTyp")
      .Store "DrCd", ADC("DmCd")
      .Store "DrSz", ADC("DmSz")
      '****** Sachin 3.02 *****
      If .IsNew(RowNum) Then
        .Store "DrDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                            " Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' And DmPrtKey='" + ctCurrPrtn + "' ")
      End If
      '****** Sachin 3.02 *****
  End With
End Sub
'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgCT
      .Store "DctTcTyp", ADC("DmTcTyp")
      .Store "DctDmCd", ADC("DmCd")
      .Store "DctDmSz", ADC("DmSz")
      '****** Sachin 3.02 *****
      If .IsNew(RowNum) Then
        .Store "DctDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                        " Where DmPrtKey='" + ctCurrPrtn + "' And DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' ")
      End If
      '****** Sachin 3.02 *****
  End With
End Sub
Private Sub GrdDsgModAvl_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgModAvl
      .Store "DmaTcTyp", ADC("DmTcTyp")
      .Store "DmaCd", ADC("DmCd")
      .Store "DmaSz", ADC("DmSz")
  End With
End Sub
Private Sub GrdDsgLab_GotFocus()
    '****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey") <> ctCurrPrtn)) Or mf_DtLocked Then
        GrdDsgLab.AllowAdd = False
        GrdDsgLab.AllowDelete = False
    Else
        GrdDsgLab.AllowAdd = True
        GrdDsgLab.AllowDelete = True
    End If
End Sub
Private Sub GrdDsgLab_RowWhen(ByVal RowNum As Integer)
  '*** If GrdDsgLab.IsNew Then DlSr = Max + 1
  With GrdDsgLab
    If .Value(RowNum, "DlSr") = 0 Then
       .Value(RowNum, "DlSr") = .MaxVal("DlSr") + 1
       .Value(RowNum, "DlQty") = 1
    End If
    If .IsNew(RowNum) Then .Value(RowNum, "DlPrtKey") = ctCurrPrtn  '****** Sachin 3.02 04-12-07 - Default value for Partition Key in RM Grid
  End With
End Sub
Private Sub GrdDsgLab_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  If ADC("DmPrtKey") <> ctCurrPrtn Then
    Cancel = True: GrdDsgLab.AllowAdd = False: GrdDsgLab.AllowDelete = False
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  
  '*** Set Helps For Each Field
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit"
    '*** Jay 2.14 *** (SEO Change)
  End If
  
  Select Case UCase(ColName)
  Case Is = UCase("DlMCd")
    Call HlpList.PMCd("LABMCD")
  Case Is = UCase("DlSCd")
    Call HlpList.PSCd("LABSCD", GrdDsgLab.Value(RowNum, "DlMCd"))
'  Case Is = UCase("DlQty")
'    If GrdDsgLab.Value(RowNum, "wQW") = "W" Then _
'       Cancel = True: ErrMsg = "Cannot Enter Labour Quantity": Exit Sub
  End Select
End Sub
Private Sub GrdDsgLab_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** wQW on DlMCd, DlSCd            Normal, Init, Copy
      'PValue('LABSCD', DlMCd, DlSCd)
  '*** DlQty on wQW                   Normal
      'If wQW<> 'Q' then 0
  
  With GrdDsgLab
    Select Case UCase(ColName)
      Case Is = UCase("wQW")
        If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then
          .Value(RowNum, "wQW") = moCn.GetFldVal("Select PValue from Param " + _
                                  "where PTyp='LABSCD' and PMCd= '" + _
                                  .Value(RowNum, "DlMCd") + "' and " + _
                                  "PSCd='" + .Value(RowNum, "DlSCd") + "'")
        End If
'      Case Is = UCase("DlQty")
'        If .Mode = fgmnorm Then
'          If .Value(RowNum, "wQW") <> "Q" Then .Value(RowNum, "DlQty") = 0
'        End If
    End Select
  End With
End Sub
Private Sub GrdDsgLab_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** DlMCd          valid PMCd('LABMCD')
  '*** DlSCd          valid PSCd('LABSCD', DlMCd) or ''
  
  With GrdDsgLab
    If .Value(RowNum, ColName) <> NewValue And UCase(ColName) <> "WQW" Then mb_Upd = True
        
    Select Case UCase(ColName)
    Case Is = UCase("DlMCd")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'LABMCD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Labour Main Code": Exit Sub
    Case Is = UCase("DlSCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                           "PTyp= 'LABSCD' And " + _
                           "PMCd='" + .Value(RowNum, "DlMCd") + "' " + _
                           "And PSCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Labour Sub Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdDsgLab_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdDsgLab
      .Store "DlTcTyp", ADC("DmTcTyp")
      .Store "DlCd", ADC("DmCd")
      .Store "DlSz", ADC("DmSz")
      '****** Sachin 3.02 *****
      If .IsNew(RowNum) Then
        .Store "DlDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                          " Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' and DmPrtKey='" + ctCurrPrtn + "' ")
      End If
      '****** Sachin 3.02 *****
  End With
End Sub
Private Sub GrdDsgAna_GotFocus()
    '****** Sachin 3.02.0 - [30/11/07] - Purge Partition - Locking Previous Partition Records
    If (ms_PwdFlag = False And ADC.AddFndMode = xFndMode) Or (UCase(ADC("DmPrtKey") <> ctCurrPrtn)) Or mf_DtLocked Then
        GrdDsgAna.AllowAdd = False
        GrdDsgAna.AllowDelete = False
    Else
        GrdDsgAna.AllowAdd = True
        GrdDsgAna.AllowDelete = True
    End If
End Sub

Private Sub GrdDsgAna_RowWhen(ByVal RowNum As Integer)
  '*** If GrdDsgAna.IsNew Then DaAnaSr = Max + 1
  With GrdDsgAna
    If .IsNew(RowNum) Then
      '.Value(RowNum, "DaAnaSr") = .MaxVal("DaAnaSr") + 1
      .Value(RowNum, "DaPrtKey") = ctCurrPrtn  '****** Sachin 3.02 04-12-07 - Default value for Partition Key in RM Grid
    End If
  End With
End Sub
Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  If ADC("DmPrtKey") <> ctCurrPrtn Then
    Cancel = True: GrdDsgAna.AllowAdd = False: GrdDsgAna.AllowDelete = False
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
  '****** Sachin 3.02.0 - Purge Partition - Previous Records to be Locked
  
  '*** Set Helps For Each Field
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    '*** Jay 2.14 *** (SEO Change)
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
    'ErrMsg = "Enter SEO Password To Edit"
    '*** Jay 2.14 *** (SEO Change)
  End If
  Select Case UCase(ColName)
  Case Is = UCase("DaAnaSr")
    Call HlpList.PMCd("DAANAFLD")
  Case Is = UCase("DaAnaCd")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAna.Value(RowNum, "DaAnaSr") + "' ") = "N" Then _
       Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
    
    Call HlpList.PSCd("DAANACD", GrdDsgAna.Value(RowNum, "DaAnaSr"))
  Case Is = UCase("DaAnaDesc")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAna.Value(RowNum, "DaAnaSr") + "' ") = "Y" Then _
       Cancel = True: ErrMsg = "Cannot Enter Desc For This Analysis Field": Exit Sub
  End Select
End Sub
Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** DaAnaCd on DaAnaSr              Normal
      'DaAnaCd= ''
  '*** DaAnaDesc on DaAnaSr            Normal
      'DaAnaDesc= ''
  '*** wAnaName on DaAnaSr             Normal
      'PDesc225('DAANAFLD', DaAnaSr)
  
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("DaAnaCd")
      If .Mode = fgmnorm Then
        .Value(RowNum, "DaAnaCd") = ""
      End If
    Case Is = UCase("DaAnaDesc")
      If .Mode = fgmnorm Then
        .Value(RowNum, "DaAnaDesc") = ""
      End If
    Case Is = UCase("wAnaName")
      If .Mode = fgmnorm Or .Mode = fgmCopy Or .Mode = fgminit Then
        .Value(RowNum, "wAnaName") = moCn.GetFldVal("Select PDesc225 from Param where " + _
                                     "PTyp= 'DAANAFLD' and PMCd= '" + .Value(RowNum, "DaAnaSr") + "' ")
      End If
    End Select
  End With
End Sub
Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** DaAnaSr        valid PMCd('DAANAFLD')
  '*** DaAnaCd        If PValue('DAANAFLD')= 'Y' then Blank or valid PSCd('DAANACD', DaAnaSr)
  '                   If PValue('DAANAFLD')= 'N' then Blank
  With GrdDsgAna
    If .Value(RowNum, ColName) <> NewValue And UCase(ColName) <> UCase("WANANAME") Then mb_Upd = True
    
    Select Case UCase(ColName)
    Case Is = UCase("DaAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
      
    Case Is = UCase("DaAnaCd")
      Dim ws_PValue As String
      ws_PValue = moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and PMCd = '" + .Value(RowNum, "DaAnaSr") + "'")
      If ws_PValue = "N" And NewValue <> "" Then Cancel = True: ErrMsg = "Analysis Code Should Not Be Entered For This Sr.": Exit Sub
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param Where " + _
                                    "PTyp= 'DAANACD' And PMCd= '" + .Value(RowNum, "DaAnaSr") + "' and " + _
                                    "PSCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Analysis Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdDsgAna_InitKey(ByVal RowNum As Integer)

  '*** Initialise Key Fields
  With GrdDsgAna
      .Store "DaTcTyp", ADC("DmTcTyp")
      .Store "DaCd", ADC("DmCd")
      .Store "DaSz", ADC("DmSz")
      '****** Sachin 3.02 *****
      If .IsNew(RowNum) Then
        .Store "DaDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                          " Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' and DmPrtKey='" + ctCurrPrtn + "' ")
      End If
      '****** Sachin 3.02 *****
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Set the RecSource for ADC
  '****** Sachin 3.02.0 - [30/11/07] - Purge Partition Key added below
  ADC.RecSource = " Select * from DsgMst " + _
                  " Where DmTcTyp ='" + ADC("DmTcTyp") + "'" + _
                  " And DMCd = '" + ADC("DMCd") + "'" + _
                  " And DMSz ='" + ADC("DMSz") + "' And DmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DmPrtKey End)"
                  
End Sub
Private Sub GrdDsgLab_SetRecSource()
  '*** Set the RecSource for GrdDsgLab
  '****** Sachin 3.02.0 - [30/11/07] - Purge Partition Key added below
  GrdDsgLab.RecSource = " Select * from DsgLab " + _
                        " Where DLTcTyp ='" + ADC("DmTcTyp") + "'" + _
                        " And DlCd = '" + ADC("DMCd") + "'" + _
                        " And DLSz ='" + ADC("DMSz") + "' And DlPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DlPrtKey End)"
End Sub
Private Sub GrdDsgRm_SetRecSource()
  '*** Set the RecSource for GrdDsgRm
  '****** Sachin 3.02.0 - [30/11/07] - Purge Partition Key added below
  GrdDsgRm.RecSource = " Select * from DsgRm " + _
                       " Where DrTcTyp ='" + ADC("DmTcTyp") + "'" + _
                       " And DrCd = '" + ADC("DMCd") + "'" + _
                       " And DrSz ='" + ADC("DMSz") + "' And DrPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DrPrtKey End)"
End Sub
Private Sub GrdDsgAna_SetRecSource()
  '*** Set the RecSource for GrdDsgAna
  '****** Sachin 3.02.0 - [30/11/07] - Purge Partition Key added below
  GrdDsgAna.RecSource = " Select * from DsgAna " + _
                        " Where DaTcTyp ='" + ADC("DmTcTyp") + "'" + _
                        " And DaCd = '" + ADC("DMCd") + "'" + _
                        " And DaSz ='" + ADC("DMSz") + "' And DaPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DaPrtKey End) Order By (case when len(DaAnaSr) = 1 then '0' else '' end)+DaAnaSr  "
End Sub
'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_SetRecSource()
  '*** Set the RecSource for GrdDsgCT
  '****** Sachin 3.02.0 - [30/11/07] - Purge Partition Key added below
  GrdDsgCT.RecSource = " Select * from DsgCT " + _
                        " Where DctTcTyp='" + ADC("DmTcTyp") + "'" + _
                        " And DctDmCd= '" + ADC("DMCd") + "'" + _
                        " And DctDmSz='" + ADC("DMSz") + "' And DctPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DctPrtKey End) Order By DctSr"
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Set the  default tab back to 0 for the tab Page TabDesgnDet
  TabDesgnDet.Tab = 0
  '****** Sachin 3.02 - Current Partition
  If ADC.MenuCd = "DM" And ADC("DMSz") = "" And ADC("DmPrtKey") = ctCurrPrtn And _
     moCn.RecSeek("Select PMCd from Param where PTyp= 'DAANAFLD' and " + _
     "PValue1= 'Y' and PMCd not in (Select DaAnaSr from DsgAna where " + _
     "DaTcTyp = '" + ADC("DmTcTyp") + "' And DaCd= '" + ADC("DMCd") + "' And " + _
     "DaSz= '" + ADC("DMSz") + "' and DaPrtKey='" + ctCurrPrtn + "') ") Then
    pr_Cancel = True: pr_ErrMsg = "Some Compulsory Analysis Field Has Not Been Entered": Exit Sub
  End If
  
  
  If moCn.RecSeek("Select 'a' from DsgRm where DrWt = 0 and " + _
     "DrTcTyp = '" + ADC("DmTcTyp") + "' And DrCd= '" + ADC("DMCd") + "' And " + _
     "DrSz= '" + ADC("DMSz") + "' and DrPrtKey='" + ctCurrPrtn + "' ") Then
    pr_Cancel = True: pr_ErrMsg = "Cannot proceed, some Rm lines have RmWt = 0": Exit Sub
  End If
  
  Dim ws_PrcsSeqArr() As String, j As Integer, ws_PrdSeqArr() As String
  Dim i As Integer
  '3.11.0
  
  If ADC("DmPrcsSeq") = "" Then
      pr_Cancel = True
      pr_ErrMsg = "Prcs Seq cannot be blank"
      Exit Sub
  End If
    
  If ADC("DmPrdSeq") <> "" Then
    ws_PrcsSeqArr = Split(ADC("DmPrcsSeq"), ",")
    ws_PrdSeqArr = Split(ADC("DmPrdSeq"), ",")
    
    If (UBound(ws_PrdSeqArr) <> UBound(ws_PrcsSeqArr)) Or (UBound(ws_PrcsSeqArr) > 0) Or (UBound(ws_PrdSeqArr) > 0) Then
      pr_Cancel = True
      pr_ErrMsg = "No of values in PrcsSeq must match that in PrdSeq and there can only be one value in PrcsSeq and PrdSeq"
      Exit Sub
    End If
  
    For i = 0 To UBound(ws_PrdSeqArr)
    If IsArray(ws_PrdSeqArr) Then
       pr_Cancel = (ADC("DmPrcsSeq") <> "") And (Not moCn.RecSeek("select vPMCd from vParam where " + _
                                        " vPTyp= 'PRDSEQ' and " + _
                                        " vPCoCd = '" + gs_CoCd + "'" + _
                                        " and vPMCd = '" + ws_PrdSeqArr(i) + "' and vPValue1='" + ws_PrcsSeqArr(i) + "'"))
       If pr_Cancel = True Then pr_ErrMsg = "Validation failed for PrdSeq, Value(" + CStr(i + 1) + ") '" + ws_PrdSeqArr(i) + " ' is invalid for given Process Sequence": Exit Sub
    End If
    Next i
  End If

  'Geeta****2.10***
  '*** Delete All The UnUsed Recs From
  moCn.Execute ("Delete From DsgTgt where " + _
                "DtTcTyp = '" + ADC("DmTcTyp") + "' And DtCd= '" + ADC("DMCd") + "' And " + _
                "DtSz= '" + ADC("DMSz") + "' ")

  If mb_Upd And moCn.RecSeek("Select 'x' From OrdDsg where OdCoCd='" + ctSelfCoCd + "'  and OdTc='PL' and  OdDmCd='" + ADC("DmCd") + "'") Then
      mb_Upd = False
      MsgBox "Design Details Updated PDM Exist. Please Check Order Zoom for PDM List."
      Exit Sub
  End If


  '*** (Jen 3.01)
  Dim wf_DiaWtDiff As Double
  If ADC("DmTotDiaWt") > 0 Then
    wf_DiaWtDiff = moCn.GetFldVal("Select convert (decimal (14, 3), " + CStr(ADC("DmTotDiaWt")) + "- " + _
                   "IsNull((Select sum(DrWt) from DsgRm where " + IIF(gs_Partition = ctCurrPrtn, " DrPrtKey='" + ctCurrPrtn + "' And ", "") + " DrTcTyp= '" + ADC("DmTcTyp") + "' " + _
                   "And DrCd= '" + ADC("DMCd") + "' And DrSz= '" + ADC("DMSz") + "' and DrRmCtg= 'D'), 0)) ")
    If wf_DiaWtDiff <> 0 Then pr_Cancel = True: pr_ErrMsg = "There Is A Difference Of " + CStr(wf_DiaWtDiff) + " In The Total Diamond Weight Specified And The BOM Weight": Exit Sub
  End If
  '*** (Jen 3.01)
  
  OLE1.Delete     ' ***** Manali 3.8.0
  
  '*** Always makes FraPicDsgZoom control as invisible***'
  FraPicDsgZoom.Visible = False

End Sub

'*** bhavna
Private Sub GrdPrcPrd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdPrcPrd
        
    Select Case UCase(ColName)
    Case Is = UCase("DPPLOC")
      Cancel = Not moCn.RecSeek("Select LocCd from Loc where " + _
                                          "LocCd = '" + NewValue + "' and LocTyp='R'")
      
      ErrMsg = "Invalid Process Location Code": Exit Sub
    Case Is = UCase("DPPPRDCTG")
    Cancel = Not moCn.RecSeek("select PMCd from Param where " + _
                             "PTyp= 'PRDCTG' and PMCd = '" + NewValue + "'")
                             
      ErrMsg = "Invalid Production Category Code": Exit Sub
    End Select
  End With

End Sub
'***********Bhavna
Private Sub GrdPrcPrd_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case UCase(ColName)
  Case Is = UCase("DPPLOC")
    Call HlpList.Loc("", "'R'", , , True)
  Case Is = UCase("DPPPRDCTG")
    Call HlpList.PMCd("PRDCTG")
  End Select

End Sub
'***********Bhavna
Private Sub GrdPrcPrd_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdPrcPrd
      .Store "DppTcTyp", ADC("DmTcTyp")
      .Store "DppDmCd", ADC("DmCd")
      If .IsNew(RowNum) Then
        .Store "DppDmIdNo", moCn.GetFldVal(" Select DmIdNo From DsgMst " + _
                                            " Where DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz='" + ADC("DmSz") + "' And DmPrtKey='" + ctCurrPrtn + "' ")
      End If

  End With

End Sub
'*****Bhavna
Private Sub GrdPrcPrd_RowWhen(ByVal RowNum As Integer)
  With GrdPrcPrd
    If .IsNew(RowNum) Then .Value(RowNum, "DppPrtKey") = ctCurrPrtn
  End With
End Sub

'***********bhavna
Private Sub GrdPrcPrd_SetRecSource()
  '*** Set the RecSource for GrdPrcPrd
  
  GrdPrcPrd.RecSource = " Select * from DsgPrd " + _
                        " Where DppTcTyp ='" + ADC("DmTcTyp") + "'" + _
                        " And DppDmCd = '" + ADC("DMCd") + "'  And DppPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DppPrtKey End)"
End Sub
'*********bhavna
Private Sub GrdPrcPrd_Validate(Cancel As Boolean)
  Cancel = GrdPrcPrd.Validate
End Sub

Private Sub GrdSumm_SetRecSource()
  '*** GrdSumm is a working grid, hence there is no recsource associated with the grid
  GrdSumm.RecSource = " "
End Sub
Private Sub ImgArr_DblClick(Index As Integer)
If ImgMatZoom.Visible = False Then
    'Dim wImgPth As String
    ImgMatZoom.Visible = True
    ImgMatZoom.Picture = ImgArr(Index).Picture
'    wImgPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , LblPic(Index).Tag, True)
'    If Dir(wImgPth) = "" Then
'        wImgPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , LblPic(Index).Tag)
'        Call ShowPic(ImgMatZoom.Height, ImgMatZoom.Width, wImgPth, ImgMatZoom)
'    End If
End If

End Sub

Private Sub ImgMatZoom_DblClick()

If ImgMatZoom.Visible = True Then _
    ImgMatZoom.Visible = False
    
End Sub

Private Sub OptFgbag_GotFocus()
  DispMsg OptFgbag.ToolTipText, etInfo
End Sub

Private Sub OptMstRec_GotFocus()
  DispMsg OptMstRec.ToolTipText, etInfo
End Sub

Private Sub TabDesgnDet_Click(PreviousTab As Integer)
  '*** Whenever a tab is selected enable the frame on the clicked tab and disable the previous tab frame
  FraDesgnDet(TabDesgnDet.Tab).Enabled = True
  FraDesgnDet(PreviousTab).Enabled = False
'Geeta ***2.10****
  If TabDesgnDet.Tab = 6 Then
    'If wb_CTMod = False Then '*** Jay 2.14Next *** [Cpy Prt Rm]
      CmdCpyRm.ZOrder (0)
      CmdCpyRm.Visible = True
    'End If                   '*** Jay 2.14Next *** [Cpy Prt Rm]
  Else
    CmdCpyRm.Visible = False
  End If
End Sub
Private Sub CmdCpy_Click()
  '*** Display or hide the Copy frame depending on the toggle state
  '*** Sachin 3.02.0 - Purge Partition - Locking Previous Partition Records
  
  If UCase(ADC("DMPRTKEY") <> ctCurrPrtn) Then DispMsg "Cannot Copy As Design In Previous Partition", etError: Exit Sub
  Call DispFra(Cpy)

  '*** Alternative Code
  '  FraCpy.Enabled = True
  '  FraNKeyAll.Enabled = False
  '  FraSumm.Enabled = False
  '  CmdCpy.Enabled = False
  '  CmdSumm.Enabled = False
  '  FraCpy.ZOrder (0)
End Sub
Private Sub CmdCpyGo_Click()
  '*** Call the Copy Routine and Close the Copy Frame
  Call CpyDsgDets
  Call CmdCpy_Click
  
  '*** Alternative Code
  '  Call CpyDsgDets
  '  FraNKeyAll.Enabled = True
  '  ADC("DmCtg").SetFocus
  '  FraCpy.Enabled = False
  '  FraSumm.Enabled = False
  '  CmdCpy.Enabled = True
  '  CmdSumm.Enabled = True
  '  FraNKeyAll.ZOrder (0)
End Sub
Private Sub CmdSumm_Click()
  '*** Display or hide the Summary frame depending on the toggle state
  '*** If The Summary Frame is visible the call the Summary Routine
  ADC("wShowPrdRmDets") = "Y"
  Call DispFra(Summ)
  If FraSumm.Visible = True Then Call Summary(ADC("wShowPrdRmDets"))

  '*** Alternative Code
  '  FraSumm.Enabled = True
  '  Call Summary
  '  FraNKeyAll.Enabled = False
  '  FraCpy.Enabled = False
  '  CmdCpy.Enabled = False
  '  CmdSumm.Enabled = False
  '  FraSumm.ZOrder (0)
End Sub
'Private Sub CmdPwd_Click()
'  Call DispFra(Pwd)
'  If ADC.AddFndMode = xFndMode Then
'    If Trim$(TxtPwd.Text) = Trim$(ms_Pwd) Then
'      ms_PwdFlag = True
'    Else
'      ms_PwdFlag = False
'    End If
'  End If
'End Sub
' ########################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ########################



'*** Jay 2.14 *** (SEO Change)
Private Sub TxtPwd_Validate(Cancel As Boolean)
  If ADC.AddFndMode = xFndMode Then
    '*** Jay 2.14 *** (SEO Change)
    If Trim(TxtPwd.text) <> "" Then
      If Trim(TxtPwd.text) = moCn.GetFldVal(" Select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   " where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
                   " UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   " (Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
        Exit Sub
      Else
        Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
      End If
    End If
    '*** Jay 2.14 *** (SEO Change)
  End If
End Sub

'Private Sub TxtPwd_Validate(Cancel As Boolean)
'JAY
'TEMP
'    If Trim(TxtPwd.text) <> "" Then
'      'If Trim$(TxtPwd.text) = Trim$(ms_Pwd) Then
'      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
'                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + ADC.MenuCd + "' and " + _
'                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
'                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
'        If UCase$(ADC("DMLockYN")) = UCase$("N") Then ms_PwdFlag = True Else ms_PwdFlag = False
'        Exit Sub
'      Else
'        ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password / User Does Not Have SEO Rights ", etError
'      End If
'    Else
'      ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
'    End If
'TEMP
'JAY
'  If ADC.AddFndMode = xFndMode Then
'    If Trim(TxtPwd.text) <> "" Then
'      If Trim$(TxtPwd.text) = Trim$(ms_Pwd) Then
'        If UCase$(ADC("DMLockYN")) = UCase$("N") Then ms_PwdFlag = True Else ms_PwdFlag = False
'        Exit Sub
'      Else
'        ms_PwdFlag = False: Cancel = True: DispMsg "Invalid Edit Password", etError
'      End If
'    Else
'      ms_PwdFlag = False
'    End If
'  End If
'End Sub
'Private Sub TxtPwd_Validate(Cancel As Boolean)
'  If ADC.AddFndMode = xFndMode And Trim(TxtPwd.Text) <> "" Then
'    If Trim(TxtPwd.Text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Edit Password", etError
'    End If
'  End If
'End Sub
' ###########################################  Manoj  ###########################################
'*** Jay 2.14 *** (SEO Change)


'*** Alternative Code
'Private Sub CmdSummGo_Click()
'  FraNKeyAll.Enabled = True
'  adc("DmCtg").SetFocus
'  FraCpy.Enabled = False
'  FraSumm.Enabled = False
'  CmdCpy.Enabled = True
'  CmdSumm.Enabled = True
'  FraNKeyAll.ZOrder (0)
'End Sub

Private Sub Pic3D_DblClick()
  '*** If wPic3DDblClk flag option is False then Enlarge the 3-Dimensional picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the 3-Dimensional picture to the original pic control size

  If ADC("DmTcTyp") = "SM" Then Exit Sub
  If wPic3DDblClk = False Then
    '*** Enlarge The Picture ***
    ' **** Zubin 211 **** '
    w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , , True)     '4.1.0.0 - Addl. Parameter Passed
    'w3DPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), False)
    
    'Sachin - 4.1.0.0
    If Dir(w3DPth) = "" Then w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"))
    
    ' **** Zubin 211 **** '
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, w3DPth, Pic3D)
    FraPic3D.Left = 0: FraPic3D.Top = 195
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
    FraPic3D.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    ' **** Zubin 211 **** '
    w3DPth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"))
    'w3DPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), False)
    ' **** Zubin 211 **** '
    Call ShowPic(wPic3DHt, wPic3DWdt, w3DPth, Pic3D)
    FraPic3D.Left = wFraPic3DLft: FraPic3D.Top = wFraPic3DTop
    FraPic3D.Height = Pic3D.Height + 100: FraPic3D.Width = Pic3D.Width + 100
  End If
  If wPic3DDblClk = False Then
    wPic3DDblClk = True
  Else
    wPic3DDblClk = False
  End If
End Sub
Private Sub PicHsk_DblClick()
  '*** If wPic3DDblClk flag option is False then Enlarge the Sketch picture to fit the FraNKeyAll Frame
  '*** If wPic3DDblClk flag option is True then bring the Sketch picture to the original pic control size
  
  If wPicHskDblClk = False Then
    '*** Enlarge The Picture ***
    ' **** Zubin 211 **** '
    wHskPth = GetPictPath(ADC("DmTcTyp"), True, ADC("DmCd"), ADC("DmSz"))
    'wHskPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), True)
    ' **** Zubin 211 **** '
     
    Call ShowLargePic(IIF(FraNKeyAll.Height - 195 > 0, FraNKeyAll.Height - 195, 0), FraNKeyAll.Width, wHskPth, PicHsk)
    FraPicHsk.Left = 0: FraPicHsk.Top = 195
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
    FraPicHsk.ZOrder (0)
  Else
    '*** Make The Picture Small ***
    ' **** Zubin 211 **** '
    wHskPth = GetPictPath(ADC("DmTcTyp"), True, ADC("DmCd"), ADC("DmSz"))
    'wHskPth = GetPictPath(ADC("DmTcTyp"), ADC("DmCd"), ADC("DmSz"), True)
    ' **** Zubin 211 **** '
    Call ShowPic(wPicHskHt, wPicHskWdt, wHskPth, PicHsk)
    FraPicHsk.Left = wFraPicHskLft: FraPicHsk.Top = wFraPicHskTop
    FraPicHsk.Height = PicHsk.Height + 100: FraPicHsk.Width = PicHsk.Width + 100
  End If
  If wPicHskDblClk = False Then
    wPicHskDblClk = True
  Else
    wPicHskDblClk = False
  End If
End Sub


Private Sub PicDsg1_DblClick()
  'When Double click the PicDsg1, It pass the pict path to PicDsgZoom Control
  'to Enlarge the PicDsg1 to show in FraNKeyAll Frame

  '*** Getting PicDsg1 path and PicDsgZoom control to be visible to avoid empty
  'images space while enlarging the picture***'
  wDsg1Pth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "1")
  PicDsgZoom.Visible = True
  
  '*** Enlarges PicDsgZoom control to the FraNKeyAll Frame and fix the properties of Height, width, etc.,***'
  Call ShowLargePic(IIF(FraNKeyAll.Height - 218 > 0, FraNKeyAll.Height - 218, 0), FraNKeyAll.Width, wDsg1Pth, PicDsgZoom)
  FraPicDsgZoom.Visible = True
  FraPicDsgZoom.Left = 0: FraPicDsgZoom.Top = 218
  FraPicDsgZoom.Height = PicDsgZoom.Height + 100: FraPicDsgZoom.Width = PicDsgZoom.Width + 100
  FraPicDsgZoom.ZOrder (0)
    
 End Sub


Private Sub PicDsg2_DblClick()
  'When Double click the PicDsg2, It pass the pict path to PicDsgZoom Control
  'to Enlarge the PicDsg2 to show in FraNKeyAll Frame
   
  '*** Getting PicDsg1 path and PicDsgZoom control to be visible to avoid empty
  'images space while enlarging the picture***'
  wDsg2Pth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "2")
  PicDsgZoom.Visible = True
  
  '*** Enlarges PicDsgZoom control to the FraNKeyAll Frame and fix the properties of Height, width, etc.,***'
  Call ShowLargePic(IIF(FraNKeyAll.Height - 218 > 0, FraNKeyAll.Height - 218, 0), FraNKeyAll.Width, wDsg2Pth, PicDsgZoom)
  FraPicDsgZoom.Visible = True
  FraPicDsgZoom.Left = 0: FraPicDsgZoom.Top = 218
  FraPicDsgZoom.Height = PicDsgZoom.Height + 100: FraPicDsgZoom.Width = PicDsgZoom.Width + 100
  FraPicDsgZoom.ZOrder (0)

End Sub


Private Sub PicDsg3_DblClick()
  'When Double click the PicDsg3, It pass the pict path to PicDsgZoom Control
  'to Enlarge the PicDsg3 to show in FraNKeyAll Frame
   
  '*** Getting PicDsg3 path and assign that path to PicDsgZoom control  ***'
  wDsg3Pth = GetPictPath(ADC("DmTcTyp"), False, ADC("DmCd"), ADC("DmSz"), , , , , , , , "3")
  PicDsgZoom.Visible = True
  
  '*** Enlarges PicDsgZoom control to the FraNKeyAll Frame and fix the properties of Height, width, etc.,***'
  Call ShowLargePic(IIF(FraNKeyAll.Height - 218 > 0, FraNKeyAll.Height - 218, 0), FraNKeyAll.Width, wDsg3Pth, PicDsgZoom)
  FraPicDsgZoom.Visible = True
  FraPicDsgZoom.Left = 0: FraPicDsgZoom.Top = 218
  FraPicDsgZoom.Height = PicDsgZoom.Height + 100: FraPicDsgZoom.Width = PicDsgZoom.Width + 100
  FraPicDsgZoom.ZOrder (0)
  
  
End Sub

Private Sub PicDsgZoom_DblClick()
  'When Double clicked the zoomed image it will become invisible
  FraPicDsgZoom.Visible = False
 
End Sub


Private Sub Summary(ByVal ms_ShowPrdRmDets As String)
  '*** Summary option is used to show the Raw Material Qty and Wt totals at the Rm Sub Ctg Level
  '*** It also shows the Gross and Category wise Totals
  '*** While calculating the Gross Wt, the Diamond & Colour Stone Wts have to be converted to Gms
  
  ' ***** Manali 3.03 -18/06/08 - Palladium included
  Dim wGrsWt As Single, wGldWt As Single, wPlWt As Single, wSilWt As Single, wPdWt As Single, wOthMtWt As Single
  Dim wDiaWt As Single, wDiaQty As Single, wCSWt As Single, wCSQty As Single
  Dim wAccWt As Single, wAccQty As Single, wGrsQty As Single, wRow As Single
  Dim wWt As String, wQty As String
  
  wGrsWt = 0: wGldWt = 0: wPlWt = 0: wOthMtWt = 0
  wDiaWt = 0: wDiaQty = 0: wCSWt = 0: wCSQty = 0
  wAccWt = 0: wAccQty = 0: wSilWt = 0: wPdWt = 0
    
    If ms_ShowPrdRmDets = "Y" Then
      wWt = "DrPrdWt": wQty = "DrPrdQty"
    ElseIf ms_ShowPrdRmDets = "N" Then
      wWt = "DrWt": wQty = "DrQty"
    End If
  '*** Open The ResultSet
  Dim wSummRs As MwfLib.MDORowSet
  '****** Sachin 3.02 04-12-07 - Summary according to the Partition Selected
  Set wSummRs = moCn.OpenRes("Select DrRmCtg, DrRmSCtg, sum(" + wQty + ") as qQty, " + _
                             "sum(" + wWt + ") as qWt " + _
                             "from DsgRm where DrTcTyp='" + ADC("DmTcTyp") + "' and " + _
                             "DrCd='" + ADC("DmCd") + "' and " + _
                             "DrSz='" + ADC("DmSz") + "' And DrPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else DrPrtKey End)" + _
                             "Group By DrRmCtg, DrRmSCtg")
  '*** Set the allow add and allow delete options of the grid to true
  GrdSumm.AllowAdd = True
  GrdSumm.AllowDelete = True
  
  '*** The DelAllRecs option is not to be used for deleting existing records in the grid instead set the dot rows property to 1
  GrdSumm.Rows = 1
  ''GrdSumm.DelAllRecs
  
  With GrdSumm
    Do While Not (wSummRs.EOF Or wSummRs.BOF)
      .AddItem
      wRow = .Rows - 1
      .Value(wRow, "wRmCtg") = wSummRs!DrRmCtg
      .Value(wRow, "wRmSCtg") = wSummRs!DrRmSCtg
      .Value(wRow, "wQty") = wSummRs!qQty
      .Value(wRow, "wWt") = wSummRs!qWt
      Select Case Trim(wSummRs!DrRmCtg)
      Case Is = "G"
        wGldWt = wGldWt + wSummRs!qWt
      Case Is = "P"
        wPlWt = wPlWt + wSummRs!qWt
      
      ' ****** sachin 2.12 ****** '
      Case Is = "S"
        wSilWt = wSilWt + wSummRs!qWt
      ' ****** sachin 2.12 ****** '
        
      ' ***** Manali 3.03 -18/06/08 - Palladium
      Case Is = "L"
        wPdWt = wPdWt + wSummRs!qWt
      ' ***** Manali 3.03 -18/06/08 - Palladium
      
      Case Is = "M"
        wOthMtWt = wOthMtWt + wSummRs!qWt
      Case Is = "D"
        'wDiaWt = wDiaWt + (wSummRs!qWt / 5)
        wDiaWt = wDiaWt + wSummRs!qWt
        wDiaQty = wDiaQty + wSummRs!qQty
      Case Is = "C"
        'wCSWt = wCSWt + (wSummRs!qWt / 5)
        wCSWt = wCSWt + wSummRs!qWt
        wCSQty = wCSQty + wSummRs!qQty
      Case Is = "X"
        wAccWt = wAccWt + wSummRs!qWt
        wAccQty = wAccQty + wSummRs!qQty
      End Select
      If Trim(wSummRs!DrRmCtg) = "C" Or Trim(wSummRs!DrRmCtg) = "D" Then
        wGrsWt = wGrsWt + (wSummRs!qWt / 5)
      Else
        wGrsWt = wGrsWt + wSummRs!qWt
      End If
      wGrsQty = wGrsQty + wSummRs!qQty
      ADC("wGrsWt") = wGrsWt: ADC("wGrsQty") = wGrsQty
      ADC("wGldWt") = wGldWt
      ADC("wPlWt") = wPlWt
      ADC("wSilWt") = wSilWt
      ADC("wPdWt") = wPdWt    ' **** Manali 3.03 - Palladium
      ADC("wOthMtWt") = wOthMtWt
      ADC("wDiaWt") = wDiaWt: ADC("wDiaQty") = wDiaQty
      ADC("wCSWt") = wCSWt: ADC("wCSQty") = wCSQty
      ADC("wAccWt") = wAccWt: ADC("wAccQty") = wAccQty
      ADC("wKt") = ADC("DmKt")
      wSummRs.MoveNext
    Loop
  End With
  
  GrdSumm.AllowAdd = False
  GrdSumm.AllowDelete = False
  Set wSummRs = Nothing
End Sub
Private Sub CpyDsgDets()

  '****** Sachin 3.02.0 - Purge Partition - The Source Designs can be only from the Current Partition

  '*** Copy Option is used to copy details from the specified Design Type, Design Cd and Size
  Dim wRow As Single
  '*** Check if copy is possible
      '(i.e. check for valid source Type-Design-Size combination)
  '*** If Source Design Type is 'DM' or 'SM' then Source Size should be ''
      'or valid PMCd('DMSZ'). If Source Dsg Type is 'PM' then Source Size should be entered
  '*** Copy is not possible from Source Dsg Type 'PM' to Target Dsg Type 'SM'
  '*** Copy is not possible if there already exists Raw Materail Records
  '*** While copying, the Design Details are copied first. If the DsgMst rec is
      'not saved then exit this procedure else continue with copying the Rm & Lab Details

  If Not moCn.RecSeek("select PMCd from Param where PTyp= 'DMTCTYP' " + _
                 " and PMCd = '" + ADC("wDmTcTyp") + "'") Then _
    DispMsg "Invalid Design Type", etError: Exit Sub

  '****** Sachin 3.02.0 - [03/12/07] - Purge Partition Partition Key added below
  If Not moCn.RecSeek("select DmCd from DsgMst where DmTcTyp= '" + ADC("wDmTcTyp") + "' " + _
                 "and DmCd = '" + ADC("wDmCd") + "' And DmPrtKey='" + ctCurrPrtn + "'") Then _
    DispMsg "Invalid Design Code", etError: Exit Sub
  
  If ADC("wDmSz") <> "" And (ADC("wDmTcTyp") = "DM" Or ADC("wDmTcTyp") = "SM") And _
     Not moCn.RecSeek("select PMCd from Param where PTyp= 'DMSZ' and " + _
                 "PMCd = '" + ADC("wDmSz") + "'") Then _
    DispMsg "Invalid Size", etError: Exit Sub
    
'*****Geeta***210
'  If ADC("wDmSz") = "" And ADC("wDmTcTyp") = "PM" Then _
'    DispMsg "Part Code Cannot Be Blank", etError: Exit Sub
  
  If (ADC("DmTcTyp") = "DM" Or ADC("DmTcTyp") = "SM") And _
     ADC("wDmTcTyp") = "PM" Then DispMsg "Cannot Copy From Component Master To Design/Sketch Master", etError: Exit Sub
  
  If ADC("DmTcTyp") = "PM" And ADC("wDmTcTyp") <> "PM" Then _
        DispMsg "Cannot Copy From Design/Sketch Master To Component Master", etError: Exit Sub
  '*** Jay 2.13(CT) ***
  
  If GrdDsgRm.Rows - 1 > 0 Then DispMsg "Cannot Copy As Raw Material Details Already Exist", etError: Exit Sub
  If GrdDsgLab.Rows - 1 > 0 Then DispMsg "Cannot Copy As Labour Details Already Exist", etError: Exit Sub
  '*** Jay 2.14Next ***
  If GrdDsgCT.Rows - 1 > 0 Then DispMsg "Cannot Copy As Component Details Already Exist", etError: Exit Sub
  '*** Jay 2.14Next ***
  '578-12
  If GrdDsgAllow.Rows - 1 > 0 Then DispMsg "Cannot Copy As Allow Details Already Exist", etError: Exit Sub
  
  '*** Beginning the copy process
  ADC.StartCopy
  MWLib.BeginProcess Me, "Copying Design Details ..."
  Dim wRsDsg As MwfLib.MDORowSet
  Dim wRsDsgRm As MwfLib.MDORowSet
  Dim wRsDsgLab As MwfLib.MDORowSet
  Dim wRsDsgAna As MwfLib.MDORowSet
  Dim wRsDsgCT As MwfLib.MDORowSet '*** Jay 2.13(CT) ***
  Dim wRsDsgAllow As MwfLib.MDORowSet
  '****** Sachin 3.02.0 - [03/12/07] - Partition Key added below
  Set wRsDsg = moCn.OpenRes("Select * from DsgMst where DmTcTyp='" + ADC("wDmTcTyp") + "' and " + _
               "DmCd='" + ADC("wDmCd") + "' and DmSz='" + ADC("wDmSz") + "' And DmPrtKey='" + ctCurrPrtn + "'")
  With wRsDsg
    If Not (.EOF Or .BOF) Then
      .MoveFirst
      ADC("DmDesc") = !DmDesc
      ADC("DmCtg") = !DmCtg
      ADC("DmPrdCtg") = !DmPrdCtg
      ADC("DmSalCtg") = !DmSalCtg
      'uni.4 SalCtg 2 and 3 added
      ADC("DmSalCtg2") = !DmSalCtg2
      ADC("DmSalCtg3") = !DmSalCtg3
      ADC("DmBagPcs") = !DmBagPcs
      ADC("DmCmCd") = !DmCmCd
      ADC("DmHld") = !DmHld
      ADC("DmHldDesc") = !DmHldDesc
      ADC("DmParts") = !DmParts
      ADC("DmPartDesc") = !DmPartDesc
      ADC("DmOldCd") = !DmOldCd
      ADC("DmCol") = !DmCol
      ADC("DmPrcsSeq") = !DmPrcsSeq   '**** uma
      ADC("DmPrdSeq") = !DmPrdSeq
      ADC("DmPrdInst") = !DmPrdInst
      ADC("DmDsgBy") = !DmDsgBy
      ADC("DmDsgDt") = !DmDsgDt
      ADC("DmModMkr") = !DmModMkr
      '$$$ System gives an error while copying (These fields are of datatype Text)
      ADC("DmHistory") = IIF(!DmHistory = "", "History", !DmHistory)
      ADC("DmFaults") = IIF(!DmFaults = "", "Faults", !DmFaults)
      ADC("DmLstMdf") = !DmLstMdf
      ADC("DmUom") = !DmUom
      ADC("DmVaCtg") = !DmVaCtg
      ADC("DmKt") = !DmKt
      ADC("DmLsCtg") = !DmLsCtg
      ADC("DmSetCd") = !DmSetCd
      
      ADC("DmRegnExc") = !DmRegnExc
      ADC("DmCmGrp") = !DmCmGrp
      ADC("DmValidYn") = "Y"        '4.3.0.0
      
      ' Zubin 212
      ADC("DmWaxWt") = !DmWaxWt
      ADC("DmModRunWt") = !DmModRunWt
      ADC("DmSilModWt") = !DmSilModWt
      ADC("DmCasPcWt") = !DmCasPcWt
      ' Zubin 212
      ADC("DmPrtKey") = !DmPrtKey       '****** Sachin 3.02.0 - Partition Key value
      If ADC("DmTcTyp") <> "PM" And ADC("DmSz") = "" Then ADC("DmDefSz") = !DmDefSz         '*** Jay 3.2.0 [DefSz]
      ' ***** Manali 3.5.0 - 08/11/08 - Source Design Code field added (Behaviour is of text- field)
      If ADC("DmTcTyp") = "DM" And ADC("wDmTcTyp") = "SM" Then
        ADC("DmSrcDsgCd") = !DmCd
      Else
        ADC("DmSrcDsgCd") = ""
      End If
      ' ***** Manali 3.5.0 - 08/11/08 - Source Design Code field added
      If ADC.SaveRec = False Then Exit Sub
      
    End If
  End With
  
  '****** Sachin 3.02.0 - Partition Key added below
  Set wRsDsgRm = moCn.OpenRes("Select * from DsgRm where DrTcTyp= '" + ADC("wDmTcTyp") + "' and " + _
             "DrCd= '" + ADC("wDmCd") + "' and DrSz= '" + ADC("wDmSz") + "' And DrPrtKey = '" + ctCurrPrtn + "' Order By DrSr")
  With wRsDsgRm
    GrdDsgRm.StartCopy
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details For Raw Material Ctg '" + !DrRmCtg + "'"
      GrdDsgRm.AddItem
      wRow = GrdDsgRm.Rows - 1
      GrdDsgRm.Value(wRow, "DrSr") = !DrSr
      GrdDsgRm.Value(wRow, "DrRmCtg") = !DrRmCtg
      GrdDsgRm.Value(wRow, "DrRmSCtg") = !DrRmSCtg
      GrdDsgRm.Value(wRow, "DrRmCd") = !DrRmCd
      GrdDsgRm.Value(wRow, "DrLn1") = !DrLn1
      GrdDsgRm.Value(wRow, "DrLn2") = !DrLn2
      GrdDsgRm.Value(wRow, "DrLn3") = !DrLn3
      GrdDsgRm.Value(wRow, "DrQty") = !DrQty
      GrdDsgRm.Value(wRow, "DrRmPtr") = !DrRmPtr
      GrdDsgRm.Value(wRow, "DrWt") = !DrWt
      GrdDsgRm.Value(wRow, "DrPrdQty") = !DrPrdQty
      GrdDsgRm.Value(wRow, "DrPrdWt") = !DrPrdWt
      GrdDsgRm.Value(wRow, "DrSetSCd") = !DrSetSCd
      GrdDsgRm.Value(wRow, "DrWsQty") = !DrWsQty
      GrdDsgRm.Value(wRow, "DrHsQty") = !DrHsQty
      GrdDsgRm.Value(wRow, "DrSubShp") = !DrSubShp
      GrdDsgRm.Value(wRow, "DrAlyCd") = !DrAlyCd
      GrdDsgRm.Value(wRow, "DrMainMet") = !DrMainMet
      GrdDsgRm.Value(wRow, "DrPrtKey") = !DrPrtKey      '****** Sachin 3.02.0 - Partition Key value
      GrdDsgRm.SaveRec (wRow)
      .MoveNext
    Loop
    GrdDsgRm.EndCopy
  End With
  
  '****** Sachin 3.02.0 - Partition Key added below
  Set wRsDsgLab = moCn.OpenRes("Select * from DsgLab where DlTcTyp='" + ADC("wDmTcTyp") + "' and " + _
             "DlCd='" + ADC("wDmCd") + "' and DlSz='" + ADC("wDmSz") + "' And DlPrtKey='" + ctCurrPrtn + "' Order By DlSr")
  With wRsDsgLab
    GrdDsgLab.StartCopy
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details For Labour Main Code '" + !DlMCd + "'"
      GrdDsgLab.AddItem
      wRow = GrdDsgLab.Rows - 1
      GrdDsgLab.Value(wRow, "DlSr") = !DlSr
      GrdDsgLab.Value(wRow, "DlMcd") = !DlMCd
      GrdDsgLab.Value(wRow, "DlScd") = !DlSCd
      GrdDsgLab.Value(wRow, "DlQty") = !DlQty
      GrdDsgLab.Value(wRow, "DlPrtKey") = !DlPrtKey    '****** Sachin 3.02.0 - Partition Key value
      GrdDsgLab.SaveRec (wRow)
      .MoveNext
    Loop
    GrdDsgLab.EndCopy
  End With
  
  '*******Geeta***2.10****************
  
 If ADC("DmSz") = "" Then
  'And ADC("DmTcTyp") = "DM" Then    '****** Sachin 3.02.0 - Partition Key added below
    'D.navin.9- Getting analysis records from source
    Set wRsDsgAna = moCn.OpenRes("Select * from DsgAna where DaTcTyp= '" + ADC("wDmTcTyp") + "' and " + _
               "DaCd= '" + ADC("wDmCd") + "' and DaSz= '' And DaPrtKey='" + ctCurrPrtn + "' Order By DaAnaSr")
    With wRsDsgAna
      GrdDsgAna.StartCopy
      Do While Not (.EOF Or .BOF)
        MWLib.RunningProcess "Copying Details For Analysis Field '" + !DaAnaSr + "'"
        GrdDsgAna.AddItem
        wRow = GrdDsgAna.Rows - 1
        GrdDsgAna.Value(wRow, "DaAnaSr") = !DaAnaSr
        GrdDsgAna.Value(wRow, "DaAnaCd") = !DaAnaCd
        GrdDsgAna.Value(wRow, "DaAnaDesc") = !DaAnaDesc
        GrdDsgAna.Value(wRow, "DaPrtKey") = ctCurrPrtn  '****** Sachin 3.02.0 - Partition Key value
        GrdDsgAna.SaveRec (wRow)
        .MoveNext
      Loop
      GrdDsgAna.EndCopy
    End With
  End If
  
  '*** Jay 2.13(CT) ***
  '****** Sachin 3.02.0 - Partition Key added below
  If ADC("wDmTcTyp") = "DM" And ADC("DmTcTyp") = "DM" And ADC("DmSz") = "" Then
    Set wRsDsgCT = moCn.OpenRes("Select * from DsgCT where DctTcTyp='DM' and " + _
               "DctDmCd='" + ADC("wDmCd") + "' and DctDmSz='' And DctPrtKey='" + ctCurrPrtn + "' Order By DctChr")
    With wRsDsgCT
      GrdDsgCT.StartCopy
      Do While Not (.EOF Or .BOF)
        MWLib.RunningProcess "Copying Details For Component Code '" + !DctChr + "'"
        GrdDsgCT.AddItem
        wRow = GrdDsgCT.Rows - 1
        GrdDsgCT.Value(wRow, "DctSr") = wRow
        GrdDsgCT.Value(wRow, "DctCd") = !DctCd
        GrdDsgCT.Value(wRow, "DctSz") = !DctSz
        GrdDsgCT.Value(wRow, "DctChr") = !DctChr
        GrdDsgCT.Value(wRow, "DctParts") = !DctParts
        GrdDsgCT.Value(wRow, "DctRmCd") = !DctRmCd
        GrdDsgCT.Value(wRow, "DctDmCol") = !DctDmCol
        GrdDsgCT.Value(wRow, "DctRmWt") = !DctRmWt
        GrdDsgCT.Value(wRow, "DctDc") = !DctDc
        GrdDsgCT.Value(wRow, "DctPrdSeq") = !DctPrdSeq
        GrdDsgCT.Value(wRow, "DctDesc") = !DctDesc
        GrdDsgCT.Value(wRow, "DctMainPrtYn") = !DctMainPrtYn
        GrdDsgCT.Value(wRow, "DctPrtKey") = ctCurrPrtn     '****** Sachin 3.02.0 - Partition Key value
        GrdDsgCT.SaveRec (wRow)
        .MoveNext
      Loop
      GrdDsgCT.EndCopy
    End With
  End If
  '*** Jay 2.13(CT) ***
  '578-12 copy allow details from source design
  If ADC("DmSz") = "" Then
    Set wRsDsgAllow = moCn.OpenRes("Select * from DmAllow where DmwDmTyp= '" + ADC("wDmTcTyp") + "' and " + _
               "DmwDmCd= '" + ADC("wDmCd") + "' and DmwDmSz= '' Order By DmwTyp")
    With wRsDsgAllow
      GrdDsgAllow.StartCopy
      Do While Not (.EOF Or .BOF)
        MWLib.RunningProcess "Copying Details For Exclusivity "
        GrdDsgAllow.AddItem
        wRow = GrdDsgAllow.Rows - 1
        GrdDsgAllow.Value(wRow, "DmwTyp") = !DmwTyp
        GrdDsgAllow.Value(wRow, "DmwCmCd") = !DmwCmCd
        GrdDsgAllow.Value(wRow, "DmwCmGrp") = !DmwCmGrp
        GrdDsgAllow.Value(wRow, "DmwReg") = !DmwReg
        GrdDsgAllow.Value(wRow, "DmwYN") = !DmwYN
        GrdDsgAllow.SaveRec (wRow)
        .MoveNext
      Loop
      GrdDsgAllow.EndCopy
    End With
  End If
  
  MWLib.EndProcess Me
  ADC.EndCopy
  Set wRsDsg = Nothing
  Set wRsDsgRm = Nothing
  Set wRsDsgLab = Nothing
  Set wRsDsgAna = Nothing
  '*** End of the copy process
End Sub
Private Sub CpyPrtRmDets()
  '*** Copy Option is used to copy details from the specified Design Type, Design Cd and Size
  Dim wRow As Single
  If GrdDsgRm.Rows - 1 > 0 Then DispMsg "Cannot Copy As Raw Material Details Already Exist", etError: Exit Sub

  MWLib.BeginProcess Me, "Copying Raw Material Details ..."
  Dim wRsDsgRm As MwfLib.MDORowSet

  '****** Sachin 2.12 12/12/05 - SJM - Silver Included Below
  '****** Sachin 3.02 03/12/07 - Designs Parts only from the current partition can be copied
  ' ****** Manali 3.03 -18/06/08 - Palladium 'L' added
  ' ****** Manali 3.03 - 25/06/08 - OldRmCd included in query
  ' ***** Manali 3.5.0 - 06/12/08 - Query changed for Change of RmCd within RmCtgs allowed for Main Metals
  ' ***** Manali 3.8.0 - Per Pc Wt
  Set wRsDsgRm = moCn.OpenRes("Select qDrRmCtg, qDrRmSCtg, qDrRmCd ,qOldDrRmCd,DrLn1,DrLn2,DrLn3,Sum(qDrQty) as qDrQty,DrRmPtr," + _
                              "Sum(qDrWt) as qDrWt,Sum(qDrPrdQty) as qDrPrdQty,Sum(qDrPrdWt) as qDrPrdWt," + _
                              "DrSetScd,Sum(qDrWsQty) qDrWsQty,Sum(qDrHsQty) as qDrHsQty,DrSubShp,DrAlyCd," + _
                              " Max(qDrMainMet)  As qDrMainMet From " + _
                              "(Select IsNull(Nrm.RmCtg, DrRmCtg) as qDrRmCtg, IsNull(Nrm.RmSCtg, DrRmSCtg) as qDrRmSCtg, IsNull(DtRmCd,DrRmCd) qDrRmCd, " + _
                              " DrRmCtg as qOldDrRmCtg, DrRmSCtg as qOldDrRmSCtg, DrRmCd as qOldDrRmCd,DrLn1,DrLn2,DrLn3,(DrQty*DctParts) as qDrQty, DrRmPtr, " + _
                              "(Case when DrRmCtg in ('G','P','S','L') and IsNull(DtRmCd,'')<>'' and IsNull(DtRmCd,'')<>DrRmCd Then " + _
                              "     (Case When IsNull(NRm.RmFndWt, 0)>0 Then IsNull(NRm.RmFndWt, 0)*DrQty Else (NRm.RmSpGrav/ORm.RmSpGrav)*DrWt End) " + _
                              "  Else DrWt End)* DctParts as qDrWt," + _
                              "(DrPrdQty*DctParts) as qDrPrdQty, " + _
                              "(Case when DrRmCtg in ('G','P','S','L') AND IsNull(DtRmCd,'')<>'' and IsNull(DtRmCd,'')<>DrRmCd Then " + _
                              "     (Case When IsNull(NRm.RmFndWt, 0)>0 Then IsNull(NRm.RmFndWt, 0)*DrQty Else (NRm.RmSpGrav/ORm.RmSpGrav)*DrPrdWt End) " + _
                              "Else DrPrdWt End) *(DctParts) as qDrPrdWt, " + _
                              "DrSetScd,(DrWsQty*DctParts) qDrWsQty,(DrHsQty*DctParts) as qDrHsQty,DrSubShp,DrAlyCd, " + _
                              "(Case when (DctMainPrtYN='Y' and DrMainMet='Y') then 'Y' else 'N' End) qDrMainMet " + _
                              " from DsgRm Join DsgCt On DrTcTyp='PM' and DrCd=DctCd and DrSz=DctSz and DctPrtKey=DrPrtKey " + _
                              " Left Outer Join DsgTgt On DrRmCtg=DtRmCtg and DrRmSCtg=DtRmSCtg  " + _
                              " Join RmMst ORm On ORm.RmCd=DrRmCd And ORm.RmPrtKey=DrPrtKey " + _
                              " Left Outer Join RmMst NRm On NRm.RmCd=DtRmCd And NRm.RmPrtKey=DrPrtKey " + _
                              "where DrPrtKey ='" + ctCurrPrtn + "' And DctTcTyp= '" + ADC("DmTcTyp") + "' and  DctDmCd= '" + ADC("DmCd") + " ' and DctDmSz= '" + ADC("DmSz") + "' ) Dr" + _
                              " Group By qDrRmCtg, qDrRmSCtg,qDrRmCd,qOldDrRmCd,DrLn1,DrLn2,DrLn3,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd " + _
                              " Order By qDrMainMet Desc, qDrRmCtg, qDrRmSCtg,qDrRmCd,DrLn1,DrLn2,DrLn3,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd ")
' Bef 3.8.0
'  Set wRsDsgRm = moCn.OpenRes("Select qDrRmCtg, qDrRmSCtg, qDrRmCd ,qOldDrRmCd,DrLn1,DrLn2,Sum(qDrQty) as qDrQty,DrRmPtr," + _
'                              "Sum(qDrWt) as qDrWt,Sum(qDrPrdQty) as qDrPrdQty,Sum(qDrPrdWt) as qDrPrdWt," + _
'                              "DrSetScd,Sum(qDrWsQty) qDrWsQty,Sum(qDrHsQty) as qDrHsQty,DrSubShp,DrAlyCd," + _
'                              " Max(qDrMainMet)  As qDrMainMet From " + _
'                              "(Select IsNull(Nrm.RmCtg, DrRmCtg) as qDrRmCtg, IsNull(Nrm.RmSCtg, DrRmSCtg) as qDrRmSCtg, IsNull(DtRmCd,DrRmCd) qDrRmCd, " + _
'                              " DrRmCtg as qOldDrRmCtg, DrRmSCtg as qOldDrRmSCtg, DrRmCd as qOldDrRmCd,DrLn1,DrLn2,(DrQty*DctParts) as qDrQty, " + _
'                              "DrRmPtr,(Case when (DrRmCtg in ('G','P','S','L') and IsNull(DtRmCd,'')<>'' " + _
'                              " and IsNull(DtRmCd,'')<>DrRmCd) then NRm.RmSpGrav/ORm.RmSpGrav else 1 End)*(DrWt*DctParts) as qDrWt," + _
'                              "(DrPrdQty*DctParts) as qDrPrdQty,(Case when (DrRmCtg in ('G','P','S','L') and IsNull(DtRmCd,'')<>''  " + _
'                              "and IsNull(DtRmCd,'')<>DrRmCd) then NRm.RmSpGrav/ORm.RmSpGrav else 1 End)*(DrPrdWt*DctParts) as qDrPrdWt, " + _
'                              "DrSetScd,(DrWsQty*DctParts) qDrWsQty,(DrHsQty*DctParts) as qDrHsQty,DrSubShp,DrAlyCd, " + _
'                              "(Case when (DctMainPrtYN='Y' and DrMainMet='Y') then 'Y' else 'N' End) qDrMainMet " + _
'                              " from DsgRm Join DsgCt On DrTcTyp='PM' and DrCd=DctCd and DrSz=DctSz and DctPrtKey=DrPrtKey " + _
'                              " Left Outer Join DsgTgt On DrRmCtg=DtRmCtg and DrRmSCtg=DtRmSCtg  " + _
'                              " Join RmMst ORm On ORm.RmCd=DrRmCd And ORm.RmPrtKey=DrPrtKey " + _
'                              " Left Outer Join RmMst NRm On NRm.RmCd=DtRmCd And NRm.RmPrtKey=DrPrtKey " + _
'                              "where DrPrtKey ='" + ctCurrPrtn + "' And DctTcTyp= '" + adc("DmTcTyp") + "' and  DctDmCd= '" + adc("DmCd") + " ' and DctDmSz= '" + adc("DmSz") + "' ) Dr" + _
'                              " Group By qDrRmCtg, qDrRmSCtg,qDrRmCd,qOldDrRmCd,DrLn1,DrLn2,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd " + _
'                              " Order By qDrMainMet Desc, qDrRmCtg, qDrRmSCtg,qDrRmCd,DrLn1,DrLn2,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd ")
  'bef 3.5.0
'  Set wRsDsgRm = moCn.OpenRes("Select DrRmCtg,DrRmSCtg,qDrRmCd,qOldDrRmCd,DrLn1,DrLn2,Sum(qDrQty) as qDrQty,DrRmPtr," + _
'                              "Sum(qDrWt) as qDrWt,Sum(qDrPrdQty) as qDrPrdQty,Sum(qDrPrdWt) as qDrPrdWt," + _
'                              "DrSetScd,Sum(qDrWsQty) qDrWsQty,Sum(qDrHsQty) as qDrHsQty,DrSubShp,DrAlyCd," + _
'                              " Max(qDrMainMet)  As qDrMainMet From " + _
'                              "(Select DrRmCtg,DrRmSCtg,IsNull(DtRmCd,DrRmCd) qDrRmCd,DrRmCd as qOldDrRmCd,DrLn1,DrLn2,(DrQty*DctParts) as qDrQty, " + _
'                              "DrRmPtr,(Case when (DrRmCtg in ('G','P','S','L') and IsNull(DtRmCd,'')<>'' " + _
'                              " and IsNull(DtRmCd,'')<>DrRmCd) then NRm.RmSpGrav/ORm.RmSpGrav else 1 End)*(DrWt*DctParts) as qDrWt," + _
'                              "(DrPrdQty*DctParts) as qDrPrdQty,(Case when (DrRmCtg in ('G','P','S','L') and IsNull(DtRmCd,'')<>''  " + _
'                              "and IsNull(DtRmCd,'')<>DrRmCd) then NRm.RmSpGrav/ORm.RmSpGrav else 1 End)*(DrPrdWt*DctParts) as qDrPrdWt, " + _
'                              "DrSetScd,(DrWsQty*DctParts) qDrWsQty,(DrHsQty*DctParts) as qDrHsQty,DrSubShp,DrAlyCd, " + _
'                              "(Case when (DctMainPrtYN='Y' and DrMainMet='Y') then 'Y' else 'N' End) qDrMainMet " + _
'                              " from DsgRm Join DsgCt On DrTcTyp='PM' and DrCd=DctCd and DrSz=DctSz and DctPrtKey=DrPrtKey " + _
'                              " Left Outer Join DsgTgt On DrRmCtg=DtRmCtg and DrRmSCtg=DtRmSCtg  " + _
'                              " Join RmMst ORm On ORm.RmCd=DrRmCd And ORm.RmPrtKey=DrPrtKey " + _
'                              " Left Outer Join RmMst NRm On NRm.RmCd=DtRmCd And NRm.RmPrtKey=DrPrtKey " + _
'                              "where DrPrtKey ='" + ctCurrPrtn + "' And DctTcTyp= '" + ADC("DmTcTyp") + "' and  DctDmCd= '" + ADC("DmCd") + " ' and DctDmSz= '" + ADC("DmSz") + "' ) Dr" + _
'                              " Group By DrRmCtg,DrRmSCtg,qDrRmCd,qOldDrRmCd,DrLn1,DrLn2,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd " + _
'                              " Order By qDrMainMet Desc,DrRmCtg,DrRmSCtg,qDrRmCd,DrLn1,DrLn2,DrRmPtr,DrSubShp,DrSetSCd, DrAlyCd ")
 
 ' **** Manali 3.5.0 - 06/12/08 - DrRmCtg replaced by qDrRmCtg, DrRmSCtg replaced by qDrRmSCtg
 With wRsDsgRm
    GrdDsgRm.StartCopy
    
    ADC("DmParts") = moCn.GetFldVal("Select Sum(DctParts) From DsgCT Where DctTcTyp='" + ADC("DmTcTyp") + "' and DctDmCd='" + ADC("DmCd") + "' And DctDmSz=''")
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Details For Raw Material Ctg '" + !qDrRmCtg + "'"
      GrdDsgRm.AddItem
      wRow = GrdDsgRm.Rows - 1
      GrdDsgRm.Value(wRow, "DrSr") = wRow
      GrdDsgRm.Value(wRow, "DrRmCtg") = !qDrRmCtg    ' !DrRmCtg
      GrdDsgRm.Value(wRow, "DrRmSCtg") = !qDrRmSCtg    ' !DrRmSCtg
      GrdDsgRm.Value(wRow, "DrRmCd") = !qDrRmCd
      GrdDsgRm.Value(wRow, "DrLn1") = !DrLn1
      GrdDsgRm.Value(wRow, "DrLn2") = !DrLn2
      GrdDsgRm.Value(wRow, "DrLn3") = !DrLn3
      GrdDsgRm.Value(wRow, "DrQty") = !qDrQty
      'GrdDsgRm.Value(wRow, "DrRmPtr") = !DrRmPtr   ' *** Manali - 3.03 - 25/06/08 - ChgPtrSTW for PrtRmCpy
      'GrdDsgRm.Value(wRow, "DrWt") = !qDrWt        ' *** Manali - 3.03 - 25/06/08 - ChgPtrSTW for PrtRmCpy
      GrdDsgRm.Value(wRow, "DrPrdQty") = !qDrPrdQty
      'GrdDsgRm.Value(wRow, "DrPrdWt") = !qDrPrdWt   ' *** Manali - 3.03 - 25/06/08 - ChgPtrSTW for PrtRmCpy
      ' *** Manali - 3.03 - 25/06/08 - ChgPtrSTW for PrtRmCpy
      Dim wRmSTWGrpOld As String, wRmDetOld As String, wOldRmSCtg As String, wOldRmCtg As String
      Dim wRmSTWGrp As String, wRmDetArr() As String, wChgPtrSTW As String
      wRmSTWGrp = "": wRmSTWGrpOld = "": wOldRmSCtg = !qDrRmSCtg    ' wOldRmSCtg = !DrRmSCtg
      wChgPtrSTW = UCase(moCn.GetFldVal("Select HChgPtrOnStwYN from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'"))

      If UCase(moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "'")) = "Y" And _
               wChgPtrSTW = "Y" And _
            (UCase(Trim(!qDrRmCtg)) = "D" Or UCase(Trim(!qDrRmCtg)) = "C") Then
            wRmDetOld = moCn.GetFldVal("Select RmCtg+','+RmSCtg+ ','+ RmSTWGrp from RmMst where RmCd= '" + !qOldDrRmCd + "' ")
            If wRmDetOld <> "" Then
                wRmDetArr = Split(wRmDetOld, ",")
                wOldRmCtg = wRmDetArr(0)
                wOldRmSCtg = wRmDetArr(1)
                wRmSTWGrpOld = wRmDetArr(2)
                wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + !qDrRmCd + "' and RmPrtKey='" + ctCurrPrtn + "' ")
            End If
            If wRmSTWGrp <> wRmSTWGrpOld Or !qDrRmSCtg <> wOldRmSCtg Then  ' !DrRmSCtg
                GrdDsgRm.Value(wRow, "DrRmPtr") = GetWtFrSz(!qDrRmCtg, !qDrRmSCtg, !DrLn1, !DrLn2, wRmSTWGrp)
            Else
                GrdDsgRm.Value(wRow, "DrRmPtr") = !DrRmPtr
            End If
      Else
            GrdDsgRm.Value(wRow, "DrRmPtr") = !DrRmPtr
      End If
      If wChgPtrSTW = "Y" And _
        (wRmSTWGrp <> wRmSTWGrpOld Or !qDrRmSCtg <> wOldRmSCtg) And _
        (UCase(Trim(!qDrRmCtg)) = "D" Or UCase(Trim(!qDrRmCtg)) = "C") Then
            GrdDsgRm.Value(wRow, "DrPrdWt") = GrdDsgRm.Value(wRow, "DrRmPtr") * !qDrPrdQty
            GrdDsgRm.Value(wRow, "DrWt") = GrdDsgRm.Value(wRow, "DrRmPtr") * !qDrPrdQty
      Else
            GrdDsgRm.Value(wRow, "DrPrdWt") = !qDrPrdWt
            GrdDsgRm.Value(wRow, "DrWt") = !qDrWt
      End If
      ' *** Manali - 3.03 - 25/06/08 - ChgPtrSTW for PrtRmCpy
      
      GrdDsgRm.Value(wRow, "DrSetSCd") = !DrSetSCd
      GrdDsgRm.Value(wRow, "DrWsQty") = !qDrWsQty
      GrdDsgRm.Value(wRow, "DrHsQty") = !qDrHsQty
      GrdDsgRm.Value(wRow, "DrSubShp") = !DrSubShp
      GrdDsgRm.Value(wRow, "DrAlyCd") = !DrAlyCd
      GrdDsgRm.Value(wRow, "DrMainMet") = !qDrMainMet
      GrdDsgRm.Value(wRow, "DrPrtKey") = ctCurrPrtn '****** Sachin 3.02.0 - Partition Key value
      GrdDsgRm.SaveRec (wRow)
      .MoveNext
    Loop
    GrdDsgRm.EndCopy
  End With
  MWLib.EndProcess Me
  Set wRsDsgRm = Nothing
  '*** End of the copy process
  
End Sub
Private Sub DispFra(ByVal pv_DsgFra As en_DsgFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  Select Case pv_DsgFra
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
      ADC("wDmTcTyp").SetFocus
      FraCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpy)
    End If
  Case Is = Summ
    If FraSumm.Visible = True Then
      FraSumm.Visible = False
      FraSumm.Enabled = False
      Call EnaDisaCmds(False, CmdSumm)
      CmdSumm.SetFocus
    Else
      FraSumm.Visible = True
      FraSumm.Enabled = True
      CmdSumm.SetFocus
      FraSumm.ZOrder
      Call EnaDisaCmds(True, CmdSumm)
    End If
  '**** Shilpa *****/
 ' Case Is = Pwd
 '   If FraPwd.Visible = True Then
 '     FraPwd.Visible = False
 '     FraPwd.Enabled = False
 '     Call EnaDisaCmds(False, CmdPwd)
 '     CmdPwd.SetFocus
 '   Else
 's     FraPwd.Visible = True
 '     FraPwd.Enabled = True
 '     CmdPwd.SetFocus
 '     FraPwd.ZOrder
 '     Call EnaDisaCmds(True, CmdPwd)
 '   End If
  '**** Shilpa *****/
   Case Is = CpyRm
    If FraPrtRmCpy.Visible = True Then
      FraPrtRmCpy.Visible = False
      FraPrtRmCpy.Enabled = False
      Call EnaDisaCmds(False, CmdCpyRm)
      CmdCpyRm.SetFocus
    Else
      FraPrtRmCpy.Visible = True
      FraPrtRmCpy.Enabled = True
   '   adc("wDmTcTyp").SetFocus
      FraPrtRmCpy.ZOrder
      Call EnaDisaCmds(True, CmdCpyRm)
    End If
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  Case Is = CpyFgBag
    If FraCpyFgBag.Visible = True Then
      FraCpyFgBag.Visible = False
      FraCpyFgBag.Enabled = False
      Call EnaDisaCmds(False, CmdCpyFgBag)
    Else
      FraCpyFgBag.Visible = True
      FraCpyFgBag.Enabled = True
      FraCpyFgBag.ZOrder
      ADC("wCoCd").SetFocus
      Call EnaDisaCmds(True, CmdCpyFgBag)
    End If
    
  '4.1.2
  Case Is = DsgMat
    If FraDsgMatrix.Visible = True Then
      FraDsgMatrix.Visible = False
      FraDsgMatrix.Enabled = False
      Call EnaDisaCmds(False, CmdDsgMat)
      CmdDsgMat.SetFocus
    Else
      FraDsgMatrix.Visible = True
      FraDsgMatrix.Enabled = True
      FraDsgMatrix.ZOrder
      Call EnaDisaCmds(True, CmdDsgMat)
    End If
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  '6.1b-46  GrdPDMLst
  Case Is = PDMLst
    If FraPDMList.Visible = True Then
      FraPDMList.Visible = False
      FraPDMList.Enabled = False
      Call EnaDisaCmds(False, CmdPDMLst)
      CmdPDMLst.SetFocus
    Else
      FraPDMList.Visible = True
      FraPDMList.Enabled = True
      CmdPDMLst.SetFocus
      FraPDMList.ZOrder
      Call EnaDisaCmds(True, CmdPDMLst)
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
      '*** Old Code
      'pv_CmdBut.BackColor = vbWhite
      'pv_CmdBut.FontBold = True
      '*** Old Code
      If pv_CmdBut.Pic = 0 Then
        pv_CmdBut.BackColor = vbWhite
        'pv_CmdBut.FontBold = True
      End If
      pv_CmdBut.CausesValidation = True
    Else
      '*** Old Code
      'pv_CmdBut.BackColor = &H8000000F
      'pv_CmdBut.FontBold = False
      '*** Old Code
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
  CmdSumm.Enabled = Not pv_ShowFra
'  CmdPwd.Enabled = Not pv_ShowFra
  CmdCpyRm.Enabled = Not pv_ShowFra
  CmdCpyFgBag.Enabled = Not pv_ShowFra    '*** Jay 2.14 *** [FG Bag Rm Copy]
  CmdOpnPdf.Enabled = Not pv_ShowFra    '***** Manali 3.8.0
  CmdMktPdf.Enabled = Not pv_ShowFra    '***** Bhavna 5.1.0
  CmdDsgMat.Enabled = Not pv_ShowFra    '*** Sachin 4.1.2
  CmdPDMLst.Enabled = Not pv_ShowFra  '6.1b-46
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
  
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCpy.Visible = False
  FraSumm.Visible = False
' FraPwd.Visible = False
  FraCpy.Enabled = False
  FraSumm.Enabled = False
' FraPwd.Enabled = False
  FraPrtRmCpy.Visible = False
  FraPrtRmCpy.Enabled = False
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  FraCpyFgBag.Visible = False
  FraCpyFgBag.Enabled = False
  '*** Jay 2.14 *** [FG Bag Rm Copy]
  FraPDMList.Visible = False    '6.1b-46
  FraPDMList.Enabled = False    '6.1b-46
  '3.11.3
  FraDsgMatrix.Visible = False
  FraDsgMatrix.Enabled = False
End Sub
'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  '*** Jay 2.13(CT) ***
  ws_BtnStr = "CmdCpyGo"
  'ws_BtnStr = "CmdCpyGo,CmdPrtRmCpyGo"
  '*** Jay 2.13(CT) ***
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
Private Sub GrdDsgLab_Validate(Cancel As Boolean)
  Cancel = GrdDsgLab.Validate
End Sub
Private Sub GrdDsgRm_Validate(Cancel As Boolean)
  Cancel = GrdDsgRm.Validate
End Sub
Private Sub GrdDsgAna_Validate(Cancel As Boolean)
  Cancel = GrdDsgAna.Validate
End Sub
'****Geeta****2.10***

'*** Jay 2.13(CT) ***
Private Sub GrdDsgCT_Validate(Cancel As Boolean)
  Cancel = GrdDsgCT.Validate
End Sub
'****** Sachin 2.13.0 - 02-06-2006 ******
Private Sub GrdDsgModAvl_Validate(Cancel As Boolean)
  Cancel = GrdDsgModAvl.Validate
End Sub
Private Sub GrdDsgTgt_Validate(Cancel As Boolean)
  Cancel = GrdDsgTgt.Validate
End Sub
Public Sub MakeRmFromFg()
  Dim wCtgStr As String, wDmCtg As String, wTmpTbl As String
  Dim wOrStr As String, wStr As String, wMainMet As Boolean
  Dim wRsRmDet As MDORowSet, wOrRecSet As MDORowSet
  Dim wDivFactor As Single, wRmRow As Single
  Dim wTotAWt As Double
  Dim wf_CnvRt As Single
  Dim wDRrTcTyp As String, wCRrTcTyp As String, wDFrOrdYN As String
  Dim wCFrOrdYN As String
  Dim wSetRt As Boolean
  Dim wTotWtDiff As Double, wWtDiff As Double, wTotIrRmWt As Double, wAddWt As Double
  Dim wIrRmWt As MDORowSet, wRmCtg As String, wFrOrdYN As String, wMainMetWt As Double
  Dim wFrRmWt As Double, wFrRmQty As Double, wOrTotWt As Double, wOrTotQty As Double
  Dim wFRmDet As MDORowSet
  Dim wSetTyp As Boolean
  Dim wOrLnStr As String
  Dim wCnd As String
  
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then DispMsg "Design Is Locked. Open Lock To Copy Details", etError: Exit Sub
  If GrdDsgRm.Rows - 1 > 0 Then GrdDsgRm.DelAllRecs

  ms_hAddJROYN = moCn.GetFldVal("Select hAddJROYN from Head where hCoCd='" + ADC("wCoCd") + "'")
    
  wCnd = ADC.RepCond
  
  Dim wRs As MDORowSet
  If OptFgbag.Value = True Then
    Set wRs = moCn.OpenResultset("select * from Fgd B Join Bag On FdBIdNo=BIdNo And FdBYy= BYy and FdBChr =BChr and FdBNo= BNo and FdPrtKey=BPrtKey " + _
                        "Where FdBYyKey=BLstYyKey And FdPrtKey='" + ctCurrPrtn + "' And " + _
                        "Not Exists(Select 'x' From Fgd M Where M.FdCoCd= B.FdCoCd and M.FdRefYy= B.FdYy and " + _
                        "M.FdRefKey= B.FdKey And M.FdPrtKey=B.FdPrtKey and (Select PValue from Param where PTyp= 'TC' and PMCd= M.FdTc)= 'FWB') " + _
                        "And (Select PValue from Param where PTyp= 'TC' and PMCd= B.FdTc)= 'FB' and " + wCnd + " ")
  ElseIf OptMstRec.Value = True Then
    Set wRs = moCn.OpenResultset("select OdTc as FdPrdOdTc, OdYy as FdPrdOdYy, OdChr as FdPrdOdChr, OdNo as FdPrdOdNo, OdSr as FdPrdOdSr, OdDmCd as FdDmCd, OdOrdQty as Fdqty, " + _
                        "OdCOCd as FdCOCd, OdTc as FdTc, OdYy as FdYy, OdChr as FdChr, OdNo as FdNo, OdSr as FdSr " + _
                        "from OrdDsg join (Select IsNull((Select OdIdNo from vwDmStkNumMaster Where MasterYN='Y' and OdDmCd='" + ADC("DmCd") + "' And " + _
                        "OdDmStkCoCd='" + ADC("wCoCd") + "' and OdDmStkYy='" + ADC("wBYy") + "' and OdDmStkChr='" + ADC("wBChr") + "' and " + _
                        "OdDmStkNo=" + CStr(ADC("wBNo")) + " ), 0) as qMstRecOdIdNo) qMstRec On qMstRecOdIdNo=OdIdNo ")
  End If
               
  Dim wRs_RmCd As MDORowSet, wb_WtFrOrd As Boolean, ws_ShowPtr As String
  Dim ws_PtrWOMax As String
  If wRs.RecCount <= 0 Then
    DispMsg "Could not find the source record to Copy from", etWarning
    Exit Sub
  End If
  wDmCtg = moCn.GetFldVal("Select DmCtg from DsgMst where DmPrtKey='" + ctCurrPrtn + "' and DmSz= '' and " + _
               "DmCd= '" + wRs!FdDmCd + "' ")
  wTmpTbl = "#TmpInvRm"
 
  Dim ws_OrLn As String, ws_Ptr As String, wCommStr As String
  ws_OrLn = "(case when Pvalue3='Y' then (Case when PValue2='Y' then  " + _
            "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty=0 then 0 else  " + _
            "cast((OrPrdWt/OrPrdQty) as Decimal(12,4)) end ) end )  " + _
            "else cast(Orln1 as Decimal(12,4)) end ) "
    
  ws_ShowPtr = moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + ctSelfCoCd + "' and HCd= '" + ctSelfCmCd + "'")
  ws_Ptr = "max(case when '" + ws_ShowPtr + "'= 'Y' then (Case when PValue2='Y' then  " + _
             "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty= 0 then 0 else  " + _
             "cast((OrPrdWt/ OrPrdQty) as Decimal(12, 4)) end ) end )  " + _
             "else 0 end) as qPtr "
    
  ws_PtrWOMax = "(case when '" + ws_ShowPtr + "'= 'Y' then (Case when PValue2='Y' then  " + _
             "cast(OrRmPtr as Decimal(12,4)) else (case when OrPrdQty= 0 then 0 else  " + _
             "cast((OrPrdWt/ OrPrdQty) as Decimal(12, 4)) end ) end )  " + _
             "else 0 end) as qPtr "
  Dim wstrForFgRm As String, wTmpTblForFgRm As String
  wTmpTblForFgRm = "#TmpInvRm2"
  wstrForFgRm = "Select ROW_NUMBER() over(order by FrRmCtg, FrRmCd) as qRunningNo,FrRmCd as qFrOldRmCd,FrRmCd as FrRmCd, FrRmSz as FrRmSz, 0 as qFrRmSz2, 0 as qFrRmSz3, sum(FrRmQty) as FrRmQty," + _
            "sum(FrRmWt) As FrRmWt,sum(FrRmWt) As FrRmAWt, max(RrRngSz) as FrRs, " + _
            "fdprtkey,max(FdCoCd) as FdCoCd,max(FdTc) as FdTc,max(FdYy) as FdYy,max(FdChr) as FdChr,max(FdNo) as FdNo,max(FdSr) as FdSr,max(FrPrtKey) as FrPrtKey,max(FrRmSz2) AS FrRmSz2,max(FrRmSz3) AS FrRmSz3 " + _
            "into #TmpInvRm2 from Fgd join FgRm on FrFdIdNo=FdIdNo And FrCoCd= FdCoCd and FrTc= FdTc and FrYy= Fdyy and FrChr= FdChr and FrNo = FdNo And FrSr = FdSr And FrPrtKey=FdPrtKey " + _
            "Join RmMst on FrRmCd = RmCd And FrPrtKey=RmPrtKey Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
           "left outer Join RmRt on RrTcTyp = 'RSZ' and RrCmCd='ZSELF' and RrCtg = FrRmCtg and RrSCtg = FrRmSCtg and FrRmSz >= RrFrLn and FrRmSz <=RrToLn " + _
            "where FdCoCd = '" + ADC("wCoCd") + "' and  FdTc= '" + wRs!FdTc + "' and FdYy= '" + wRs!FdYy + "' and FdChr= '" + wRs!FdChr + "' " + _
            "and FdNo =" + CStr(wRs!FdNo) + " and FdSr =" + CStr(wRs!FdSr) + " And FdPrtKey='C' " + _
            "Group by FrRmCtg, FrRmCd, FrRmSz ,fdprtkey "
            
  moCn.CreateTmpTable wTmpTblForFgRm, wstrForFgRm
  Dim wStr1 As String
  Dim wStr2 As String
  Dim wStr3 As String
  Dim wstrJoinCnd As String
  Dim wstrToExecute As String
  If OptFgbag.Value = True Then
    wStr = "Select qRunningNo, 1123.234 as qOrTotRmWt, 0 as qOrTotRmQty, (case when (isnull(qOrRmCd, ''))= '' then '' else 'Y' end) as qOrRmExists, FrRmCd as qFrOldRmCd," + _
           "FrRmCd as qFrRmCd, FrRmSz as qFrRmSz, 0 as qFrRmSz2, 0 as qFrRmSz3, " + _
           "(FrRmQty)/" + CStr(wRs!Fdqty) + " as qFrRmQty, (FrRmWt)/" + CStr(wRs!Fdqty) + " As qFrRmAWt, (Isnull(qOrWt,0)) as qOrWt, (IsNull(qOrQty,0)) as qOrQty, " + _
           " IsNull(qOrSetSCd, '') as qOrSetSCd, " + _
           "(IsNull(qOrMainMet,'')) as qFrMainMet, (RmCtg) as qRmCtg, (RmSCtg) as qRmSCtg, qOrRs, " + _
           "IsNull(qOrLn2,0) as qFrLn2,IsNull(qOrLn3,0) as qFrLn3, (IsNull(qOrWsQty,0)) as qFrWsQty, (IsNull(qOrHsQty,0)) as qFrHsQty," + _
           "(IsNull(qOrSubShp,'')) as qFrSubShp, (IsNull(qOrAlyCd,''))as qFrAlyCd, " + _
           "(PValue3) as qStkByPtr, " + _
           "(IsNull(qPtr, (case when PValue3= 'Y' then FrRmSz else 0 end) )) as qPtr, " + _
           "(IsNull(qLn1, (case when PValue3= 'N' then FrRmSz else 0 end) )) as qLn1, " + _
           "(IsNull(qLn2, (case when PValue3= 'N' then FrRmSz2 else 0 end) )) as qLn2, " + _
           "(IsNull(qLn3, (case when PValue3= 'N' then FrRmSz3 else 0 end) )) as qLn3, " + _
           "(RmSTWGrp) as qRmSTWGrp "
     wStr1 = " into           " + wTmpTbl + " "
     wstrJoinCnd = " qOrRmCd= FrRmCd and (qOrLn1= FrRmSz)  "
     wStr2 = " from " + wTmpTblForFgRm + " " + _
           "Join RmMst on FrRmCd = RmCd And FrPrtKey=RmPrtKey Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
           "Join (select space(12) as qOrRs, OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
           "(OrMainMet) as qOrMainMet, (OrWt) as qOrWt, (OrQty) as qOrQty, OrLn2 as qOrLn2,OrLn3 as qOrLn3, (OrWsQty) as qOrWsQty, " + _
           "(OrHsQty) as qOrHsQty, (OrSubShp) as qOrSubShp, (OrAlyCd) as qOrAlyCd, " + _
           ws_PtrWOMax + ", (OrLn1) as qLn1,  (OrLn2) as qLn2, (OrLn3) as qLn3 from OrdRm " + _
           "Join RmMst on RmCd= OrRmCd And RmPrtKey=OrPrtKey " + _
           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
           "where OrCoCd= '" + ADC("wCoCd") + "' /* OrCoCd= '" + gs_CoCd + "' */ and " + _
           " OrTc = '" + wRs!FdPrdOdTc + "' and OrYy= '" + wRs!FdPrdOdYy + "' " + _
           " and OrChr= '" + wRs!FdPrdOdChr + "' and OrNo = " + CStr(wRs!FdPrdOdNo) + _
           " and OrRmCtg in ('D','C') and OrSr= " + CStr(wRs!FdPrdOdSr) + " And OrPrtKey='" + ctCurrPrtn + "' " + _
           ") Or2  On " + wstrJoinCnd + _
           "where FdCoCd = '" + ADC("wCoCd") + "' /* FdCoCd = '" + gs_CoCd + "' */ and " + _
           " FdTc= '" + wRs!FdTc + "' and FdYy= '" + wRs!FdYy + "' and FdChr= '" + wRs!FdChr + "' and " + _
           "FdNo = " + CStr(wRs!FdNo) + " and FdSr = " + CStr(wRs!FdSr) + " And FdPrtKey='" + ctCurrPrtn + "'" + _
           ""
    wstrToExecute = wStr + wStr1 + wStr2
    moCn.CreateTmpTable wTmpTbl, wstrToExecute
    
    wstrJoinCnd = " qOrRmCd= FrRmCd and (qOrRs = FrRS)  "
    wStr1 = ""
    wStr2 = " from " + wTmpTblForFgRm + " " + wStr1 + _
           "Join RmMst on FrRmCd = RmCd And FrPrtKey=RmPrtKey Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
           "Join (select RrRngSz as qOrRs, OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
           "(OrMainMet) as qOrMainMet, (OrWt) as qOrWt, (OrQty) as qOrQty, OrLn2 as qOrLn2,OrLn3 as qOrLn3, (OrWsQty) as qOrWsQty, " + _
           "(OrHsQty) as qOrHsQty, (OrSubShp) as qOrSubShp, (OrAlyCd) as qOrAlyCd, " + _
           ws_PtrWOMax + ", (OrLn1) as qLn1,  (OrLn2) as qLn2, (OrLn3) as qLn3 from OrdRm " + _
           "Join RmMst on RmCd= OrRmCd And RmPrtKey=OrPrtKey " + _
           "Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
           "Join RmRt on RrTcTyp = 'RSZ' and RrCmCd='ZSELF' and RrCtg = OrRmCtg and RrSCtg = OrRmSCtg and OrLn1 >= RrFrLn and OrLn1<=RrToLn " + _
           "where OrCoCd= '" + ADC("wCoCd") + "' /* OrCoCd= '" + gs_CoCd + "' */ and " + _
           " OrTc = '" + wRs!FdPrdOdTc + "' and OrYy= '" + wRs!FdPrdOdYy + "' " + _
           " and OrChr= '" + wRs!FdPrdOdChr + "' and OrNo = " + CStr(wRs!FdPrdOdNo) + _
           " and OrSr= " + CStr(wRs!FdPrdOdSr) + " And OrPrtKey='" + ctCurrPrtn + "' " + _
           ") Or2  On " + wstrJoinCnd + _
           "where not exists (Select 'a' from #TmpInvRm where #TmpInvRm.qFrRmCd = #TmpInvRm2.FrRmCd and #TmpInvRm.qFrRmSz = #TmpInvRm2.FrRmSz) and FdCoCd = '" + ADC("wCoCd") + "' /* FdCoCd = '" + gs_CoCd + "' */ and " + _
           " FdTc= '" + wRs!FdTc + "' and FdYy= '" + wRs!FdYy + "' and FdChr= '" + wRs!FdChr + "' and " + _
           "FdNo = " + CStr(wRs!FdNo) + " and FdSr = " + CStr(wRs!FdSr) + " And FdPrtKey='" + ctCurrPrtn + "'" + _
           ""
    
    wstrToExecute = "Insert into " + wTmpTbl + " " + wStr + wStr2
    moCn.Execute (wstrToExecute)
    wstrJoinCnd = " qOrRmCd= FrRmCd "
    wStr1 = ""
    wStr2 = " from " + wTmpTblForFgRm + " " + wStr1 + _
           "Join RmMst on FrRmCd = RmCd And FrPrtKey=RmPrtKey Join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
           "Join (select space(12) as qOrRs, OrRmCd as qOrRmCd, " + ws_OrLn + " as qOrLn1, OrSetSCd as qOrSetSCd, " + _
           "(OrMainMet) as qOrMainMet, (OrWt) as qOrWt, (OrQty) as qOrQty, OrLn2 as qOrLn2,OrLn3 as qOrLn3, (OrWsQty) as qOrWsQty, " + _
           "(OrHsQty) as qOrHsQty, (OrSubShp) as qOrSubShp, (OrAlyCd) as qOrAlyCd, " + _
           ws_PtrWOMax + ", (OrLn1) as qLn1,  (OrLn2) as qLn2, (OrLn3) as qLn3 from OrdRm " + _
           " Join RmMst on RmCd= OrRmCd And RmPrtKey=OrPrtKey " + _
           " Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
           " where OrCoCd= '" + ADC("wCoCd") + "' /* OrCoCd= '" + gs_CoCd + "' */ and " + _
           " OrTc = '" + wRs!FdPrdOdTc + "' and OrYy= '" + wRs!FdPrdOdYy + "' " + _
           " and OrChr= '" + wRs!FdPrdOdChr + "' and OrNo = " + CStr(wRs!FdPrdOdNo) + _
           " and OrSr= " + CStr(wRs!FdPrdOdSr) + " And OrPrtKey='" + ctCurrPrtn + "' " + _
           ") Or2  On " + wstrJoinCnd + _
           "where not exists (Select 'a' from #TmpInvRm where #TmpInvRm.qFrRmCd = #TmpInvRm2.FrRmCd and #TmpInvRm.qFrRmSz = #TmpInvRm2.FrRmSz) and FdCoCd = '" + ADC("wCoCd") + "' /* FdCoCd = '" + gs_CoCd + "' */ and " + _
           " FdTc= '" + wRs!FdTc + "' and FdYy= '" + wRs!FdYy + "' and FdChr= '" + wRs!FdChr + "' and " + _
           "FdNo = " + CStr(wRs!FdNo) + " and FdSr = " + CStr(wRs!FdSr) + " And FdPrtKey='" + ctCurrPrtn + "'" + _
           ""
    
    wstrToExecute = "Insert into " + wTmpTbl + " " + wStr + wStr2
    moCn.Execute (wstrToExecute)
    moCn.Execute ("Update #TmpInvRm" + _
                      " set qOrTotRmWt=OrTotRmWt," + _
                      " qOrTotRmQty = OrTotRmQty " + _
                      "    From" + _
                      " ( " + _
                      "  Select qRunningNo, sum(qOrWt) as OrTotRmWt, sum(qOrQty) as OrTotRmQty from " + _
                      "   #TmpInvRm group by qRunningNo " + _
                      " ) ForTotals " + _
                      " where #TmpInvRm.qRunningNo = ForTotals.qRunningNo " + _
                      "  ")

    Dim qRunningNo As Long
    Dim qCurrentRunningNo As Long
    Dim wFrTotWt As Double
    Dim wFrTotQty As Long
    qCurrentRunningNo = 0
    Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)
    
    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
          If wRsRmDet!qRunningNo <> qCurrentRunningNo Then
            qCurrentRunningNo = wRsRmDet!qRunningNo
            wFrTotQty = wRsRmDet!qFrRmQty
            wFrTotWt = wRsRmDet!qFrRmAWt
          End If
          
          If wRsRmDet!qRmCtg = "D" Or wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "G" Or wRsRmDet!qRmCtg = "P" Or wRsRmDet!qRmCtg = "S" Or wRsRmDet!qRmCtg = "L" Or wRsRmDet!qRmCtg = "M" Or wRsRmDet!qRmCtg = "X" Then
            wRsRmDet.edit
            wRsRmDet!qFrRmAWt = Round(MWLib.Div(wRsRmDet!qOrWt, wRsRmDet!qOrTotRmWt) * wRsRmDet!qFrRmAWt, 3)
            wRsRmDet!qFrRmQty = Round(MWLib.Div(wRsRmDet!qOrQty, wRsRmDet!qOrTotRmQty) * wRsRmDet!qFrRmQty)
            wFrTotWt = wFrTotWt - wRsRmDet!qFrRmAWt
            wFrTotQty = wFrTotQty - wRsRmDet!qFrRmQty
            wRsRmDet.Update
        End If
        wRsRmDet.MoveNext
    Loop
    Set wFRmDet = Nothing

  ElseIf OptMstRec.Value = True Then
    wCommStr = " Select (case when isnull(OrRmCd, '')= '' then '' else 'Y' end) as qOrRmExists, OrRmCd as qFrOldRmCd,OrRmCd as qFrRmCd, " + _
           "OrLn1 as qFrRmSz, 0 as qFrRmSz2, 0 as qFrRmSz3, (OrQty) as qFrRmQty, (OrWt) As qFrRmAWt, (OrWt) as qOrWt," + _
           "OrMainMet AS qFrMainMet, (OrQty) as qOrQty,  IsNull(OrSetSCd, '') as qOrSetSCd,(RmCtg) as qRmCtg," + _
           "(RmSCtg) as qRmSCtg, IsNull(OrLn2,0) as qFrLn2,IsNull(OrLn3,0) as qFrLn3, (IsNull(OrWsQty,0)) as qFrWsQty," + _
           "(IsNull(OrHsQty,0)) as qFrHsQty,(IsNull(OrSubShp,'')) as qFrSubShp, (IsNull(OrAlyCd,''))as qFrAlyCd," + _
           "(PValue3) as qStkByPtr, (case when (PValue3)= 'Y' then (OrRmPtr) else 1 end)  as qPtr, " + _
           "(case when (PValue3)= 'N' then OrLn1 else 0 end)  as qLn1 ,(case when (PValue3)= 'N' then OrLn2 else 0 end)  as qLn2," + _
           "(case when (PValue3)= 'N' then OrLn3 else 0 end)  as qLn3, (RmSTWGrp) as qRmSTWGrp "
    wStr = wCommStr + "into " + wTmpTbl + " From OrdRm " + _
           "join (Select IsNull((Select OdIdNo from vwDmStkNumMaster Where OdDmCd='" + ADC("DmCd") + "' And " + _
           "OdDmStkCoCd='" + ADC("wCoCd") + "' and MasterYN='Y' and OdDmStkYy='" + ADC("wBYy") + "' and OdDmStkChr='" + ADC("wBChr") + "' and " + _
           "OdDmStkNo=" + CStr(ADC("wBNo")) + " ), 0) as qMstRecOdIdNo) qMstRec On (qMstRecOdIdNo=OrOdIdNo or ('Y'='" + ms_hAddJROYN + "' and OrOdMstIdNo=qMstRecOdIdNo)) " + _
           "join OrdDsg on OdIdNo =qMstRecOdIdNo " + _
           "join RmMst on OrRmCd = RmCd And OrPrtKey =RmPrtKey Join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
           "where OdDmCd='" + ADC("DmCd") + "' And OdDmStkCoCd='" + ADC("wCoCd") + "' and OdDmStkYy='" + _
           ADC("wBYy") + "' and OdDmStkChr='" + ADC("wBChr") + "' and  OdDmStkNo=" + CStr(ADC("wBNo")) + _
           " "
    moCn.CreateTmpTable wTmpTbl, wStr
  End If

  Dim ws_Join As String, ws_Cnd As String
  Set wRsRmDet = moCn.OpenResultset("Select * from " + wTmpTbl)

  If (wRsRmDet.RecCount > 0) Then wRsRmDet.MoveFirst
  
  Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
    wRsRmDet.edit
    Dim qTgtRmCd As String
    Select Case ADC("wCpyFgBagGrp")
    Case Is = "-"
      ws_Join = ""
      ws_Cnd = ""
    Case Is = "RMCOL"
      ws_Join = "Join RmMst on RmCd ='" + wRsRmDet!qFrRmCd + "' "
      ws_Cnd = " and DtCol=RmCol"
    Case Is = "RMCOL-SZ"
      ws_Join = "Join RmMst on RmCd ='" + wRsRmDet!qFrRmCd + "' " + _
              "Left Outer Join RmRt on RrTcTyp ='RSZ' and RrCtg='" + wRsRmDet!qRmCtg + "' and " + _
              "RrSCtg='" + wRsRmDet!qRmSCtg + "' and (" + CStr(wRsRmDet!qFrRmSz) + ">=RrFrLn and " + CStr(wRsRmDet!qFrRmSz) + "<=RrToLn) "
      ws_Cnd = " and DtCol=RmCol and IsNull(RrRngSz,'') = DtRmRSz "
    Case Is = "RMRNGSZ"
      ws_Join = "Left Outer Join RmRt on RrTcTyp ='RSZ' and RrCtg='" + wRsRmDet!qRmCtg + "' and " + _
              "RrSCtg='" + wRsRmDet!qRmSCtg + "' and (" + CStr(wRsRmDet!qFrRmSz) + ">=RrFrLn and " + CStr(wRsRmDet!qFrRmSz) + "<=RrToLn) "
      ws_Cnd = " and IsNull(RrRngSz,'') = DtRmRSz "
    End Select
    qTgtRmCd = moCn.GetFldVal("Select IsNull(DtRmCd,'') From DsgTgt " + _
                    ws_Join + " where DtTcTyp='" + ADC("DmTcTyp") + "' and DtCd='" + ADC("DmCd") + "' and DtSz='" + ADC("DmSz") + "' and " + _
                    "DtRmCtg='" + wRsRmDet!qRmCtg + "' and DtRmSCtg='" + wRsRmDet!qRmSCtg + "'" + ws_Cnd)
              
    If qTgtRmCd <> "" Then wRsRmDet!qFrRmCd = qTgtRmCd
  
    wRsRmDet.Update
    wRsRmDet.MoveNext
  Loop
  
  Set wRsRmDet = moCn.OpenRes("select * from " + wTmpTbl + " Order By qFrMainMet")
  wTotAWt = 0
  With GrdDsgRm
    .StartCopy
    Do While Not (wRsRmDet.EOF Or wRsRmDet.BOF)
      .AddItem
      wRmRow = .Rows - 1
        wDivFactor = IIF(wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D", 5, 1)
        .Value(wRmRow, "DrSr") = wRmRow
        .Value(wRmRow, "DrRmCtg") = wRsRmDet!qRmCtg
        .Value(wRmRow, "DrRmSCtg") = wRsRmDet!qRmSCtg
        .Value(wRmRow, "DrRmCd") = wRsRmDet!qFrRmCd

        .Value(wRmRow, "DrLn1") = wRsRmDet!qLn1
        .Value(wRmRow, "DrRmPtr") = wRsRmDet!qPtr
        
        If .Value(wRmRow, "DrRmPtr") = 0 Then
          .Value(wRmRow, "DrRmPtr") = GetWtFrSz(wRsRmDet!qRmCtg, wRsRmDet!qRmSCtg, .Value(wRmRow, "DrLn1"), .Value(wRmRow, "DrLn2"), wRsRmDet!qRmSTWGrp)
        End If
    
        .Value(wRmRow, "DrLn1") = wRsRmDet!qFrRmSz
        .Value(wRmRow, "DrLn2") = wRsRmDet!qFrLn2
        .Value(wRmRow, "DrLn3") = wRsRmDet!qFrLn3
        Dim wRmSTWGrpOld As String, wRmSTWGrp As String, wOldRmCtg As String, wOldRmSCtg As String
        Dim wRmDetOld As String, wRmDetArr() As String
        wRmSTWGrp = "": wRmSTWGrpOld = "": wOldRmSCtg = wRsRmDet!qRmSCtg
        
        If UCase(moCn.GetFldVal("Select HPtrYN from Head where HCoCd= '" + ADC("wCoCd") + "' and HCd='" + ctSelfCmCd + "'")) = "Y" And _
              (ChkFCSTW.Value = vbChecked) And (UCase(Trim(wRsRmDet!qRmCtg)) = "D" Or UCase(Trim(wRsRmDet!qRmCtg)) = "C") Then

          wRmDetOld = moCn.GetFldVal("Select RmCtg+','+RmSCtg+ ','+ RmSTWGrp from RmMst where RmCd= '" + wRsRmDet!qFrOldRmCd + "' ")
              
          If wRmDetOld <> "" Then
            wRmDetArr = Split(wRmDetOld, ",")
            wOldRmCtg = wRmDetArr(0)
            wOldRmSCtg = wRmDetArr(1)
            wRmSTWGrpOld = wRmDetArr(2)
            wRmSTWGrp = moCn.GetFldVal("Select RmSTWGrp from RmMst where RmCd= '" + wRsRmDet!qFrRmCd + "' ")
          End If
          If wRmSTWGrp <> wRmSTWGrpOld Or wRsRmDet!qRmSCtg <> wOldRmSCtg Then
            .Value(wRmRow, "DrRmPtr") = GetWtFrSz(wRsRmDet!qRmCtg, wRsRmDet!qRmSCtg, .Value(wRmRow, "DrLn1"), .Value(wRmRow, "DrLn2"), wRmSTWGrp)
          End If
        End If
        
        If wRs!Fdqty > 0 Then
          .Value(wRmRow, "DrQty") = Round((wRsRmDet!qFrRmQty / 1), 0)
          
          .Value(wRmRow, "DrWt") = (wRsRmDet!qFrRmAWt / 1)
          .Value(wRmRow, "DrPrdQty") = Round((wRsRmDet!qFrRmQty / 1), 0)
          .Value(wRmRow, "DrPrdWt") = (wRsRmDet!qFrRmAWt / 1)
          
          If wRsRmDet!qRmCtg = "C" Or wRsRmDet!qRmCtg = "D" Then
            If False And .Value(wRmRow, "DrRmPtr") <> 0 Then     '3.11.0
                .Value(wRmRow, "DrWt") = (.Value(wRmRow, "DrRmPtr") * .Value(wRmRow, "DrQty"))
                .Value(wRmRow, "DrPrdWt") = (.Value(wRmRow, "DrRmPtr") * .Value(wRmRow, "DrPrdQty"))
            End If
            .Value(wRmRow, "DrWsQty") = .Value(wRmRow, "DrQty") * (wRsRmDet!qFrWsQty / (IIF((wRsRmDet!qFrWsQty + wRsRmDet!qFrHsQty) > 0, (wRsRmDet!qFrWsQty + wRsRmDet!qFrHsQty), 1)))
            .Value(wRmRow, "DrHsQty") = .Value(wRmRow, "DrQty") - .Value(wRmRow, "DrWsQty")
          End If
        End If
        If ((ChkFCSTW.Value = vbChecked) Or (wRmSTWGrp <> wRmSTWGrpOld Or wRsRmDet!qRmSCtg <> wOldRmSCtg)) And _
            (UCase(Trim(wRsRmDet!qRmCtg)) = "D" Or UCase(Trim(wRsRmDet!qRmCtg)) = "C") Then
              .Value(wRmRow, "DrWt") = .Value(wRmRow, "DrRmPtr") * .Value(wRmRow, "DrQty")
              .Value(wRmRow, "DrPrdWt") = .Value(wRmRow, "DrRmPtr") * .Value(wRmRow, "DrPrdQty")
        End If

        .Value(wRmRow, "DrSetSCd") = wRsRmDet!qOrSetSCd
        .Value(wRmRow, "DrAlyCd") = wRsRmDet!qFrAlyCd
        .Value(wRmRow, "DrSubShp") = wRsRmDet!qFrSubShp
        .Value(wRmRow, "DrMainMet") = wRsRmDet!qFrMainMet
        .Value(wRmRow, "DrPrtKey") = ctCurrPrtn     '****** Sachin 3.02.0 - Setting the Partition Key value
      .SaveRec (wRmRow)
      wRsRmDet.MoveNext
    Loop
    .EndCopy
  End With
  Set wRsRmDet = Nothing
  Set wOrRecSet = Nothing
End Sub

Sub PushToCurrent()
    '****** Sachin 3.02.0 - [03/12/07] - Purge Partition - Pushing Design To Current
    On Error GoTo errhld
    Dim wRs_Ord As MDORowSet, wOldPrtKey As String
    Dim ws_TblNm As String, ws_QryStr As String, wrsRrTcTyp As MDORowSet
  
      'If moCn.RecSeek(" Select 'x' From RmMst Rm " + _
                      " Left Outer Join RmMst B On B.RmCd=Rm.RmBaseCd and Rm.RmBaseCd<>'' " + _
                      " Where Rm.RmZ='N' and Rm.RmCd='" + GrdRm.Value(GrdRm.Row, "RmCd") + "' And Rm.RmValidYN='N'") Then _
      'DispMsg "Cannot Push An InValid RmCode", etError: Exit Sub
  
      'Temporary Table Names
      ws_TblNm = "#Tmp" + ADC.MenuCd
      wOldPrtKey = ADC("DmPrtKey")
      
      ADC("DmPrtKey") = ctCurrPrtn
      If ADC.SaveRec = False Then GoTo errhld
      
      MWLib.BeginProcess Me, "Pushing Raw Material Code(s) ..."
      moCn.BeginTrans (mdoCommitRec)
      
      'Creating Temporary Table
      ws_QryStr = "Select Rm.RmCd as qRmCd,Rm.RmZCd qRmZCd, " + _
                    " Rm.RmBaseCd as qRmBaseCd,isNull(B.RmZCd,'') qRmBaseZCd " + _
                    "   Into " + ws_TblNm + _
                    "   From RmMst Rm " + _
                    " Left Outer Join RmMst B On B.RmCd=Rm.RmBaseCd and Rm.RmBaseCd<>'' " + _
                    " Where Rm.RmZ='N' and " + _
                    "   (Rm.RmCd in (Select DrRmCd From DsgRm Where DrTcTyp='" + ADC("DmTcTyp") + "' and DrCd='" + ADC("DmCd") + "') " + _
                    "                Or Rm.RmCd in (Select OrRmCd From OrdRm " + _
                    "                                  Join OrdDsg On OrPrtKey=OdPrtKey and OrOdIdNo=OdIdNo and OrCoCd=OdCoCd and OrTc=OdTc " + _
                    "                                       and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr " + _
                    "  Where OdDmCd='" + ADC("DmCd") + "' and (Select PValue From  Param where PTyp='TC' and PMCd=OdTc)='PL') " + _
                    " Or Rm.RmCd in (Select DrAlyCd From DsgRm where DrTcTyp='" + ADC("DmTcTyp") + "' and DrCd='" + ADC("DmCd") + "')) "
        
        moCn.CreateTmpTable ws_TblNm, ws_QryStr
  
        '****** Pushing All designs (with or without size) of Current Design Code
        MWLib.RunningProcess "Pushing Designs of Other Sizes..."
        moCn.Execute ("Update DsgMst Set DmPrtKey='" + ctCurrPrtn + "' " + _
                       " Where DmPrtKey<>'" + ctCurrPrtn + "' And DmTcTyp='" + ADC("DmTcTyp") + "' And DmCd='" + ADC("DmCd") + "' And DmSz<>'" + ADC("DmSz") + "'")

        '****** Pushing DsgRm
        MWLib.RunningProcess "Pushing Raw Materials ..."
        moCn.Execute ("Update DsgRm Set DrPrtKey='" + ctCurrPrtn + "' " + _
                      "From DsgRm Where DrTcTyp='" + ADC("DmTcTyp") + "' and DrCd='" + ADC("DmCd") + "' and DrPrtKey<>'" + ctCurrPrtn + "' ")
    
        '****** Pushing Raw Material Master of Corresponding RM's in DsgRm
        moCn.Execute (" Update RmMst Set RmPrtKey='" + ctCurrPrtn + "' " + _
                      " From " + ws_TblNm + " Rm where RmPrtKey<>'" + ctCurrPrtn + "' and (RmCd in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        Set wrsRrTcTyp = moCn.OpenRes(" Select PMCd From Param Where PTyp='RRTCTYP' ")
        wrsRrTcTyp.MoveFirst
        Do While Not (wrsRrTcTyp.BOF Or wrsRrTcTyp.EOF)
          moCn.Execute (" Update RmRt Set ModUsr='" + gs_UsrCd + "', RrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                        " Where RrTcTyp='" + wrsRrTcTyp!PMCd + "' And RrPrtKey<>'" + ctCurrPrtn + "' and (RrCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")
          
          moCn.Execute (" Update RmRtHist Set ModUsr='" + gs_UsrCd + "', RhPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                        " where RhTcTyp='" + wrsRrTcTyp!PMCd + "' And RhPrtKey<>'" + ctCurrPrtn + "' and (RhCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")
          wrsRrTcTyp.MoveNext
        Loop

        'moCn.Execute (" Update RmRt Set RrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                      " Where RrPrtKey<>'" + ctCurrPrtn + "' and (RrCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")
        
        'moCn.Execute (" Update RmRtHist Set RhPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                      " where RhPrtKey<>'" + ctCurrPrtn + "' and (RhCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        moCn.Execute (" Update CustRm Set CrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                      " Where CrPrtKey<>'" + ctCurrPrtn + "' and (CrOurRmCd in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        '****** Pushing Customer Raw Material Code(s) of above RM's
        moCn.Execute ("Update CustRm Set CrPrtKey='" + ctCurrPrtn + "' " + _
                      " Where CrPrtKey<>'" + ctCurrPrtn + "' and CrOurRmCd In (Select DrRmCd From DsgRm " + _
                      " Where DrTcTyp='" + ADC("DmTcTyp") + "' and drcd='" + ADC("DmCd") + "') ")
                  
        MWLib.RunningProcess "Pushing Design(s) Labour ..."
        '****** Pushing DsgLab
        moCn.Execute ("Update DsgLab Set DlPrtKey='" + ctCurrPrtn + "' " + _
                        "From DsgLab Where DlTcTyp='" + ADC("DmTcTyp") + "' and DlCd='" + ADC("DmCd") + "' and DlPrtKey<>'" + ctCurrPrtn + "' ")

        MWLib.RunningProcess "Pushing Design(s) Analysis ..."
        '****** Pushing DsgAna
        moCn.Execute ("Update DsgAna Set DaPrtKey='" + ctCurrPrtn + "' " + _
                      "From DsgLab Where DaTcTyp='" + ADC("DmTcTyp") + "' and DaCd='" + ADC("DmCd") + "' and DaPrtKey<>'" + ctCurrPrtn + "' ")

        MWLib.RunningProcess "Pushing Design(s) Component Records ..."
        '****** Pushing DsgCT
        moCn.Execute ("Update DsgCT Set DctPrtKey='" + ctCurrPrtn + "' " + _
                      "From DsgCT Where DctTcTyp='" + ADC("DmTcTyp") + "' and DctCd='" + ADC("DmCd") + "' and DctPrtKey<>'" + ctCurrPrtn + "' ")

        '****** Pushing the Corresponding Design In Customer Design Table
        MWLib.RunningProcess "Pushing Customer Design(s) Records ..."
        moCn.Execute ("Update CustDsg Set CdPrtKey='" + ctCurrPrtn + "' " + _
                            " From CustDsg  " + _
                            " Where CdOurDmCd='" + ADC("DmCd") + "' and CdPrtKey<>'" + ctCurrPrtn + "'")
                            
        '****** Pushing the Corresponding Design In Production Points Table
        MWLib.RunningProcess "Pushing Production Points Design(s) Records ..."
        moCn.Execute ("Update PrdPts Set PpPrtKey='" + ctCurrPrtn + "' " + _
                            " Where PpDmCd='" + ADC("DmCd") + "' and PpPrtKey<>'" + ctCurrPrtn + "'")

    If ADC("DmTcTyp") = "DM" Then
          '****** Pushing the OrdMst of the PDC Entry
          MWLib.RunningProcess "Pushing Design(s) PDC Records ..."
          moCn.Execute (" Update OrdDsg Set OdPrtKey='" + ctCurrPrtn + "' Where (Select PValue From  Param where PTyp='TC' and PMCd=OdTc)='PL' " + _
                        " And OdPrtKey<>'" + ctCurrPrtn + "' And OdDmCd='" + ADC("DmCd") + "'")
          
          '****** Pushing the OrdRM of the PDC Entry
          MWLib.RunningProcess "Pushing Design(s) PDC Raw Material Records ..."
          moCn.Execute ("Update OrdRm Set OrPrtKey='" + ctCurrPrtn + "' From OrdDsg where OdIdNo=OrOdIdNo and OdCoCd=OrCoCd And OdTc=OrTc And OdYy=OrYy " + _
                        " And OdChr=OrChr And OdNo=OrNo And OdSr=OrSr and OrPrtKey <>'" + ctCurrPrtn + "' and (Select PValue From  Param where PTyp='TC' and PMCd=OdTc)='PL' And OdDmCd='" + ADC("Dmcd") + "'")

          '****** Pushing the OrdLab of the PDC Entry
          MWLib.RunningProcess "Pushing Design(s) PDC Labour Records ..."
          moCn.Execute (" Update OrdLab Set OlPrtKey='" + ctCurrPrtn + "'  From OrdDsg " + _
                        " Where OdIdNo=OlOdIdNo and OdCoCd=OlCoCd And OdTc=OlTc And OdYy=OlYy And OdChr=OlChr And OdNo=OlNo And OdSr=OlSr " + _
                        "       and OlPrtKey<>'" + ctCurrPrtn + "' and (Select PValue From  Param where PTyp='TC' and PMCd=OdTc)='PL' And OdDmCd='" + ADC("DmCd") + "'")
          
          '****** Pushing the OrdCt of the PDC Entry
          MWLib.RunningProcess "Pushing Design(s) PDC Component Records ..."
          moCn.Execute (" Update OrdCT Set OctPrtKey='" + ctCurrPrtn + "' From OrdDsg " + _
                        " Where OdIdNo=OctOdIdNo and OdCoCd=OctCoCd And OdTc=OctTc And OdYy=OctYy And OdChr=OctChr And OdNo=OctNo And OdSr=OctSr " + _
                        " and OctPrtKey<>'" + ctCurrPrtn + "' and (Select PValue From  Param where PTyp='TC' and PMCd=OdTc)='PL' And OdDmCd='" + ctCurrPrtn + "'")

    End If
                        
    moCn.CommitTrans (mdoCommitRec)
    
    '****** Refreshing Grids ******
    GrdDsgRm.Init
    GrdDsgRm.RefreshData
    GrdDsgLab.Init
    GrdDsgLab.RefreshData
    GrdDsgCT.Init
    GrdDsgCT.RefreshData
    GrdDsgAna.Init
    GrdDsgAna.RefreshData
    
    '****** Ending the Process ******
    CmdPushCurr.Enabled = False
    MWLib.EndProcess Me
    Exit Sub
    
errhld:
    ADC("DmPrtKey") = wOldPrtKey
    ADC.SaveRec
    moCn.RollbackTrans mdoCommitNone
    MsgBox Err.Description & Chr(13) & " Could Not Push The RmCode to Current"
    MWLib.EndProcess Me
    '****** Sachin 3.02.0 - [03/12/07] - Purge Partition - Pushing Design To Current
    
End Sub

'****** Sachin 3.3.0 (07-07-08) - [Region Exclusivity for Multiple Customers] ******
Private Function MultiRegnCust(ByVal ps_DmCmcd As String) As Boolean

    Dim ws_DmCmCdArr() As String, CustCtr As Integer, wCustRegnCd As String
    
    MultiRegnCust = False
    ws_DmCmCdArr = Split(ps_DmCmcd, ",")
    
    If IsArray(ws_DmCmCdArr) Then
      For CustCtr = 0 To UBound(ws_DmCmCdArr)
        If wCustRegnCd <> "" And wCustRegnCd <> moCn.GetFldVal("Select CmRegnCd From CustMst Where CmCtg='C' and CmCd = '" + ws_DmCmCdArr(CustCtr) + "'") Then MultiRegnCust = True: Exit For
        wCustRegnCd = moCn.GetFldVal("Select CmRegnCd From CustMst Where CmCtg='C' and CmCd = '" + ws_DmCmCdArr(CustCtr) + "'")
      Next CustCtr
    End If

End Function
'****** Sachin 3.3.0 (07-07-08) - [Region Exclusivity for Multiple Customers] ******
'*** Jay 3.4[DP]
Private Sub GrdDsgPrm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
'*** Jay 3.4[DP]
 With GrdDsgPrm
    Select Case UCase(ColName)
    Case Is = UCase("DpTyp")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                           "PTyp= 'DSGPRM' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Design Parameter Type": Exit Sub
    Case Is = UCase("DpCd")
      If NewValue <> "" Or GrdDsgPrm.Value(RowNum, "DpTyp") = "COLL" Then
        Cancel = Not moCn.RecSeek("Select DpCd From DsgPrm Where " + _
                             "DpTyp= '" + GrdDsgPrm.Value(RowNum, "DpTyp") + "' And DpCd = '" + NewValue + "'")
        ErrMsg = "Invalid Design Parameter Code": Exit Sub
      End If
    End Select
  End With
'*** Jay 3.4[DP]
End Sub

'*** Jay 3.4[DP]
Private Sub GrdDsgPrm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
'*** Jay 3.4[DP]
 If ADC("DmPrtKey") <> ctCurrPrtn Then
    Cancel = True: GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
 End If
  
 If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
 End If
  
  '*** Set Helps For Each Field
  Select Case UCase(ColName)
  Case Is = UCase("DpTyp")
    If Not GrdDsgPrm.IsNew(RowNum) Then Cancel = True: ErrMsg = "Can Not Change Type"
    Call HlpList.PMCd("DSGPRM")
  Case Is = UCase("DpCd")
    If GrdDsgPrm.Value(RowNum, "DpTyp") = "COLL" Then
      Call HlpList.DsgPrmCd(GrdDsgPrm.Value(RowNum, "DpTyp"))
    Else
      Call HlpList.DsgPrmCd(GrdDsgPrm.Value(RowNum, "DpTyp"), True)
    End If
  End Select
'*** Jay 3.4[DP]
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrm_InitKey(ByVal RowNum As Integer)
'*** Jay 3.4[DP]
  With GrdDsgPrm
    If .IsNew(RowNum) Then
      .Store "DpDmCd", ADC("DmCd")
      .Store "DpNo", 0
      .Store "DpDesc", ""
      .Store "DpDt", MWLib.EmptyDate
      .Store "DpPctNm", ""
      .Store "DpRem1", ""
      .Store "DpRem2", ""
      .Store "DpDmNo", 0
    End If
  End With
End Sub

Private Sub GrdDsgPrm_RowValid(ByVal RowNum As Integer, Cancel As Boolean, ErrMsg As String)
'*** Jay 3.4[DP]
  With GrdDsgPrm
    If .IsNew(RowNum) And (UCase(.Value(GrdDsgPrm.Row, "DpTyp")) <> "" And UCase(.Value(GrdDsgPrm.Row, "DpTyp")) <> "COLL") And (.Value(GrdDsgPrm.Row, "DpCd") = "0" Or .Value(GrdDsgPrm.Row, "DpCd") = "") Then
      .Value(GrdDsgPrm.Row, "DpCd") = moCn.AutoGen("DsgPrm", "DpCd", .Value(GrdDsgPrm.Row, "DpTyp"), -1)
    End If
  End With
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrm_SetRecSource()
'*** Jay 3.4[DP]
    If ADC("DmTcTyp") = "DM" Then
        GrdDsgPrm.RecSource = " Select * from DsgPrm " + _
                           " Where DpDmCd ='" + ADC("DMCd") + "' and DpTyp <>'CAT' Order By DpTyp, DpCd"
    Else
        GrdDsgPrm.RecSource = " Select * from DsgPrm where 1= 2 "
    End If
End Sub

Private Sub GrdDsgPrmDet_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
'*** Jay 3.4[DP]
  If ADC("DmPrtKey") <> ctCurrPrtn Then
    Cancel = True: GrdDsgPrm.AllowAdd = False: GrdDsgPrm.AllowDelete = False
    ErrMsg = "Designs In Previous Partition Cannot Be Edited": Exit Sub
  End If
  If ms_PwdFlag = False And ADC.AddFndMode = xFndMode Then
    Cancel = True
    ErrMsg = "Design Is Locked Open Lock To Edit": Exit Sub
  End If
  If UCase(GrdDsgPrm.Value(GrdDsgPrm.Row, "DpTyp")) = "COLL" Then
    Cancel = True: ErrMsg = "Collection Can Not Be Edited": Exit Sub
  End If
  
  Select Case UCase(ColName)
  Case Is = UCase("DpDmCd")
    Call HlpList.DmCd(ADC("DmTcTyp"), , , "C")
  End Select
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_LostFocus()
'*** Jay 3.4[DP]
  Call ShowPic(wPicDPHt, wPicDPWdt, "", PicDP)
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_RowWhen(ByVal RowNum As Integer)
'*** Jay 3.4[DP]
  Dim wDpPth As String
  If UCase(GrdDsgPrm.Value(GrdDsgPrm.Row, "DpTyp")) <> "COLL" Then
    GrdDsgPrmDet.AllowAdd = True
    GrdDsgPrmDet.AllowDelete = True
  Else
    GrdDsgPrmDet.AllowAdd = False
    GrdDsgPrmDet.AllowDelete = False
  End If
  
  wDpPth = GetPictPath("DM", False, GrdDsgPrmDet.Value(RowNum, "DpDmCd"))
  Call ShowPic(wPicDPHt, wPicDPWdt, wDpPth, PicDP)
'*** Jay 3.4[DP]
End Sub
'*** Jay 3.4[DP]

Private Sub GrdDsgPrmDet_SetRecSource()
'*** Jay 3.4[DP]
    GrdDsgPrmDet.RecSource = " Select * from DsgPrm " + _
                           " Where DpTyp ='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpTyp") + "' " + _
                           " And DpCd ='" + GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd") + "' " + _
                           " And DpDmCd not In ('','" + ADC("DmCd") + "') Order By DpDmCd"
End Sub
Private Sub GrdDsgPrmDet_InitKey(ByVal RowNum As Integer)
'*** Jay 3.4[DP]
  With GrdDsgPrmDet
    If .IsNew(RowNum) Then
      .Store "DpTyp", GrdDsgPrm.Value(GrdDsgPrm.Row, "DpTyp")
      .Store "DpCd", GrdDsgPrm.Value(GrdDsgPrm.Row, "DpCd")
      .Store "DpNo", 0
      .Store "DpDesc", ""
      .Store "DpDt", MWLib.EmptyDate
      .Store "DpPctNm", ""
      .Store "DpRem1", ""
      .Store "DpRem2", ""
      .Store "DpDmNo", 0
    End If
  End With
End Sub
Private Sub GrdDsgPrm_Validate(Cancel As Boolean)
'*** Jay 3.4[DP]
  Cancel = GrdDsgPrm.Validate
End Sub
Private Sub GrdDsgPrmDet_Validate(Cancel As Boolean)
'*** Jay 3.4[DP]
  Cancel = GrdDsgPrmDet.Validate
End Sub
'6.1b-46 adding PDCm list to the grid
Public Sub ShowPDMLst()
  Dim wRsPDMLst As MDORowSet, i As Integer, wVchrStr As String
  Set wRsPDMLst = moCn.OpenRes("Select top(99) OdTc ,OdYy ,OdChr ,OdNo,OdSr ,OdOmCmCd ,OdDmCd ,OdSfx ,OdDmSz From OrdDsg " + _
                          " where OdCoCd='" + ctSelfCoCd + "'  and OdTc='PL' and  OdDmCd='" + ADC("DmCd") + "'" + _
                          " order by OdTc ,OdYy ,OdChr ,OdNo,OdSr")
  GrdPDMLst.Rows = 1
  With GrdPDMLst
    Do While Not (wRsPDMLst.EOF Or wRsPDMLst.BOF)
      .AddItem
      i = .Rows - 1
      wVchrStr = wRsPDMLst!OdTc + "/" + wRsPDMLst!OdYy + "/" + wRsPDMLst!OdChr + "/" + CStr(wRsPDMLst!OdNo) + "/" + CStr(wRsPDMLst!OdSr)
      .Value(i, "WPDMVCHR") = wVchrStr
      .Value(i, "WPDMcmctg") = "C"
      .Value(i, "WPDMCmcd") = wRsPDMLst!OdOmCmCd
      .Value(i, "WPDMDmcd") = wRsPDMLst!OdDmCd
      .Value(i, "WPDMsfx") = wRsPDMLst!OdSfx
      .Value(i, "WPDMsz") = wRsPDMLst!OdDmSz
      wRsPDMLst.MoveNext
    Loop
  End With
  Set wRsPDMLst = Nothing
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyH And Shift = 2 Then
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
  End If
  
End Sub

