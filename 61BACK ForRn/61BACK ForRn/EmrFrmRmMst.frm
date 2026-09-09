VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlforspeed8.ocx"
Begin VB.Form EmrFrmRmMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Master"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdPushCurr 
      Height          =   435
      Left            =   6750
      TabIndex        =   35
      Top             =   9330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Push To &Curr."
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
      Left            =   7005
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9405
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   30
      TabIndex        =   19
      Top             =   9330
      Width           =   14490
      _ExtentX        =   25559
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   -36
      TabIndex        =   21
      Top             =   30
      Width           =   15165
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   2090
         TabIndex        =   1
         ToolTipText     =   "Enter Raw Material Sub Category From"
         Top             =   -15
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRMSCTGFR"
         CmpStr          =   "RmSCtg >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   2090
         TabIndex        =   2
         ToolTipText     =   "Enter Raw Material Sub Category To"
         Top             =   270
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WRMSCTGTO"
         CmpStr          =   "RmSCtg <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   3975
         TabIndex        =   3
         ToolTipText     =   "Enter Raw Material Code From "
         Top             =   -15
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WRMCDFR"
         CmpStr          =   "RmCd >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   1
         Left            =   465
         TabIndex        =   0
         ToolTipText     =   "Enter Raw Material Category"
         Top             =   -15
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   2
         DataType        =   4
         IdName          =   "WRMCTG"
         CmpStr          =   "RmCtg = "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   3975
         TabIndex        =   4
         ToolTipText     =   "Enter Raw Material Code To"
         Top             =   270
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         MaxLength       =   16
         DataType        =   4
         IdName          =   "WRMCDTO"
         CmpStr          =   "RmCd <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   13635
         TabIndex        =   9
         ToolTipText     =   "Enter Partition Key (C-Current,P-Previous,Blank For All)"
         Top             =   -15
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WRMPRTKEY"
         CmpStr          =   "RmPrtKey = "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   7395
         TabIndex        =   5
         ToolTipText     =   "Enter Customer Code From "
         Top             =   -15
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WCMCDFR"
         CmpStr          =   "RmCmCd >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   9195
         TabIndex        =   6
         ToolTipText     =   "Enter Customer Code To"
         Top             =   -15
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WCMCDTO"
         CmpStr          =   "RmCmCd <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   12195
         TabIndex        =   8
         ToolTipText     =   "Enter 'Y' to view Only Customer Records 'N' for All"
         Top             =   -15
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         DataType        =   4
         IdName          =   "WCUSTYN"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   13635
         TabIndex        =   10
         ToolTipText     =   "Sort Records On"
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WRMSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   11
         Left            =   7395
         TabIndex        =   7
         ToolTipText     =   "Enter Customer Code Selection"
         Top             =   270
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   503
         MaxLength       =   255
         DataType        =   4
         IdName          =   "WCMCDSEL"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CmCd Sel"
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
         Left            =   6405
         TabIndex        =   40
         Top             =   270
         Width           =   1005
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
         Index           =   8
         Left            =   8880
         TabIndex        =   39
         Top             =   0
         Width           =   405
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "CmCd  Fr"
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
         Left            =   6405
         TabIndex        =   38
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust (Y/N)"
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
         Left            =   11160
         TabIndex        =   37
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label LblPrtnKey 
         BackStyle       =   0  'Transparent
         Caption         =   "Prtn.(C/P)"
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
         Left            =   12720
         TabIndex        =   36
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort On"
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
         Left            =   12720
         TabIndex        =   28
         Top             =   270
         Width           =   885
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "           To"
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
         Left            =   3040
         TabIndex        =   27
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "RmCd  Fr"
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
         Left            =   3040
         TabIndex        =   26
         Top             =   -15
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "             To"
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
         Left            =   1050
         TabIndex        =   25
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "SubCtg  Fr"
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
         Left            =   1050
         TabIndex        =   24
         Top             =   -15
         Width           =   1215
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
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   -15
         Width           =   495
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8910
      Left            =   -120
      TabIndex        =   20
      Top             =   480
      Width           =   15345
      Begin VB.Frame FraRmDesc 
         Height          =   1155
         Left            =   130
         TabIndex        =   29
         Top             =   7590
         Width           =   14490
         Begin MwfCtl.MWCTL_BTN1 CmdDescOk 
            Height          =   525
            Left            =   13635
            TabIndex        =   17
            Top             =   510
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
         Begin MwfCtl.MWCTL_MED TXTDESC 
            Height          =   285
            Index           =   0
            Left            =   1830
            TabIndex        =   13
            ToolTipText     =   "Enter Raw Material Description"
            Top             =   180
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   503
            MaxLength       =   90
            ReCalcParent    =   "WRMPRTDESC"
            IdName          =   "WRMDESC"
         End
         Begin MwfCtl.MWCTL_MED TXTDESC 
            Height          =   285
            Index           =   1
            Left            =   1830
            TabIndex        =   14
            ToolTipText     =   "Enter International Quality Code"
            Top             =   465
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            MaxLength       =   12
            DataType        =   4
            IdName          =   "WRMINTQLY"
         End
         Begin MwfCtl.MWCTL_MED TXTDESC 
            Height          =   285
            Index           =   2
            Left            =   1830
            TabIndex        =   16
            ToolTipText     =   "Enter Raw Material Print Description"
            Top             =   750
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   503
            MaxLength       =   90
            ReCalcOn        =   "WRMDESC"
            IdName          =   "WRMPRTDESC"
         End
         Begin MwfCtl.MWCTL_MED TXTDESC 
            Height          =   285
            Index           =   3
            Left            =   13110
            TabIndex        =   15
            ToolTipText     =   "Enter Valid (Y/N)"
            Top             =   465
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WRMVALIDYN"
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
            Left            =   11820
            TabIndex        =   34
            Top             =   465
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Print Description"
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
            Left            =   90
            TabIndex        =   33
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "International Qly"
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
            Left            =   90
            TabIndex        =   31
            Top             =   465
            Width           =   1755
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            TabIndex        =   30
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   150
         TabIndex        =   23
         Top             =   -390
         Width           =   15150
         Begin MwfCtl.MWCTL_FLX GrdRm 
            Height          =   4860
            Left            =   0
            TabIndex        =   11
            Top             =   510
            Width           =   15090
            _ExtentX        =   26617
            _ExtentY        =   8573
            Cols            =   29
            colname1        =   "RMSCTG"
            heading1        =   "Sub Ctg"
            datafld1        =   "RmSCtg"
            datatype1       =   4
            maxlength1      =   5
            tooltiptext1    =   "Enter Raw Material Sub Category"
            colname2        =   "RMCD"
            heading2        =   "Rm Code                                  "
            datafld2        =   "RmCd"
            datatype2       =   4
            maxlength2      =   16
            tooltiptext2    =   "Enter Raw Material Code"
            colname3        =   "RMUOM"
            heading3        =   "Uom     "
            datafld3        =   "RmUom"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Enter Unit Of Measurement"
            colname4        =   "RMKT"
            heading4        =   "Karat   "
            datafld4        =   "RmKt"
            datatype4       =   4
            maxlength4      =   5
            tooltiptext4    =   "Enter Karatage"
            colname5        =   "RMCOL"
            heading5        =   "Colour "
            datafld5        =   "RmCol"
            datatype5       =   4
            maxlength5      =   5
            tooltiptext5    =   "Enter Raw Material Colour"
            colname6        =   "RMQW"
            heading6        =   "Q/W Sal"
            datafld6        =   "RmQw"
            datatype6       =   4
            recalcparent6   =   "GrdRm(RMQWCST)"
            maxlength6      =   1
            style6          =   0
            tooltiptext6    =   "Enter Whether Sale Rate to be by Q/W."
            colname7        =   "RMQWCST"
            heading7        =   "Q/W Cst"
            datafld7        =   "RmQwCst"
            datatype7       =   4
            recalcon7       =   "GrdRm(RMQW)"
            maxlength7      =   1
            tooltiptext7    =   "Enter Whether Cost Rate to be by Q/W."
            colname8        =   "RMSPGRAV"
            heading8        =   "SpGravity"
            datafld8        =   "RmSpGrav"
            datatype8       =   2
            mask8           =   "##0.0000"
            maxlength8      =   8
            tooltiptext8    =   "Enter Specific Gravity"
            colname9        =   "RMPURITYRT"
            heading9        =   "Rt Purity "
            datafld9        =   "RmPurityRt"
            datatype9       =   2
            mask9           =   "##0.0000"
            maxlength9      =   8
            tooltiptext9    =   "Enter Raw Material Purity"
            colname10       =   "RMPURITYWT"
            heading10       =   "PrdPurity"
            datafld10       =   "RmPurityWt"
            datatype10      =   2
            mask10          =   "##0.0000"
            maxlength10     =   8
            tooltiptext10   =   "Enter Production Purity"
            colname11       =   "RMPURITYZ"
            heading11       =   "ZPurity     "
            datafld11       =   "RmPurityZ"
            datatype11      =   2
            mask11          =   "##0.0000"
            maxlength11     =   8
            tooltiptext11   =   "Enter Customs Purity"
            colname12       =   "RMSTWGRP"
            heading12       =   "STW Grp"
            datafld12       =   "RmSTWGrp"
            datatype12      =   4
            maxlength12     =   5
            tooltiptext12   =   "Enter Size To Weight Group"
            colname13       =   "RMFNDWT"
            heading13       =   "Fnd Wt/Pc"
            datafld13       =   "RmFndWt"
            datatype13      =   2
            mask13          =   "##0.000"
            maxlength13     =   7
            style13         =   0
            tooltiptext13   =   "Enter Per Piece Finding Wt"
            colname14       =   "RMCMCD"
            heading14       =   "Customer   "
            datafld14       =   "RmCmCd"
            datatype14      =   4
            maxlength14     =   8
            tooltiptext14   =   "Enter Customer Code"
            colname15       =   "RMZCD"
            heading15       =   "Customs Cd                             "
            datafld15       =   "RmZCd"
            datatype15      =   4
            maxlength15     =   16
            tooltiptext15   =   "Enter Customs Code"
            colname16       =   "RMDESC"
            datafld16       =   "RmDesc"
            maxlength16     =   90
            style16         =   0
            tooltiptext16   =   "Enter Description"
            colname17       =   "RMINTQLY"
            datafld17       =   "RmIntQly"
            datatype17      =   4
            maxlength17     =   12
            style17         =   0
            colname18       =   "RMPRTDESC"
            datafld18       =   "RmPrtDesc"
            maxlength18     =   90
            style18         =   0
            colname19       =   "RMZ"
            datafld19       =   "RmZ"
            datatype19      =   4
            maxlength19     =   1
            style19         =   0
            colname20       =   "RMBASECD"
            heading20       =   "Base Code                               "
            datafld20       =   "RmBaseCd"
            datatype20      =   4
            recalcparent20  =   "GrdRm(RMBASEPURITYWT)"
            maxlength20     =   16
            tooltiptext20   =   "Enter Base Code"
            colname21       =   "RMBASEPURITYWT"
            heading21       =   "BasePrdPurity"
            datafld21       =   "RmBasePurityWt"
            datatype21      =   2
            mask21          =   "##0.0000"
            recalcon21      =   "GrdRm(RMBASECD)"
            maxlength21     =   8
            tooltiptext21   =   "Enter Base Production Purity"
            colname22       =   "RMID"
            heading22       =   "RM Id            "
            datafld22       =   "RmId"
            datatype22      =   4
            maxlength22     =   8
            tooltiptext22   =   "Enter RM Id"
            colname23       =   "RMPARENTCD"
            heading23       =   "RM Parent Code"
            datafld23       =   "RmParentCd"
            datatype23      =   4
            maxlength23     =   8
            tooltiptext23   =   "Enter Parent Code"
            colname24       =   "RMALYCD"
            heading24       =   "Aly Cd                                      "
            datafld24       =   "RmAlyCd"
            datatype24      =   4
            maxlength24     =   16
            colname25       =   "RMSTNCLS"
            heading25       =   "Stone Class   "
            datafld25       =   "RmStnCls"
            datatype25      =   4
            maxlength25     =   8
            tooltiptext25   =   "Enter Stone Class"
            colname26       =   "WDESC"
            colname27       =   "RMVALIDYN"
            datafld27       =   "RmValidYN"
            maxlength27     =   1
            style27         =   0
            tooltiptext27   =   "Enter Valid (Y/N)"
            colname28       =   "RMPRTKEY"
            heading28       =   "Prtn "
            datafld28       =   "RmPrtKey"
            maxlength28     =   1
            style28         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdRmAlyCmp 
            Height          =   2310
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Enter Alloy Composition Details Here"
            Top             =   5400
            Width           =   5340
            _ExtentX        =   9419
            _ExtentY        =   4075
            Cols            =   5
            colname1        =   "RACMPCTG"
            heading1        =   "Ctg  "
            datafld1        =   "RaCmpCtg"
            datatype1       =   4
            maxlength1      =   2
            tooltiptext1    =   "Enter Composite Raw Material Category"
            colname2        =   "RACMPSCTG"
            heading2        =   "Sub Ctg"
            datafld2        =   "RaCmpSCtg"
            datatype2       =   4
            maxlength2      =   5
            tooltiptext2    =   "Enter Composite Raw Material Sub Category"
            colname3        =   "RACMPCD"
            heading3        =   "Cmp Rm Cd                             "
            datafld3        =   "RaCmpCd"
            datatype3       =   4
            maxlength3      =   16
            tooltiptext3    =   "Enter Composite Raw Material Code"
            colname4        =   "RAPERGMWT"
            heading4        =   "Wt Per Gm"
            datafld4        =   "RaPerGmWt"
            datatype4       =   2
            mask4           =   "0.000"
            recalcparent4   =   "WRAPERGMWT"
            maxlength4      =   5
            tooltiptext4    =   "Enter Weight Per Gram Of Composite Rm Cd In The Alloy"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   3090
            TabIndex        =   32
            ToolTipText     =   "Total Weight Per Gram"
            Top             =   7710
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "0.000"
            MaxLength       =   5
            DataType        =   2
            ReCalcOn        =   "GrdRmAlyCmp(RAPERGMWT)"
            IdName          =   "WRAPERGMWT"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmRmMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--- general comments on the Entry
' 2 Menus : Rm Master, Customs Master. field RmZ = 'Y' or 'N', passed from Menu.
'     No difference, except in RmZ='N', RmZCd should be valid RmCd of RmZ='Y', in Same Ctg.
' working field wRmSrt decides the Sort order for the Entry. (from PDesc225 of 'SRTRM')
' nonKey When sets the Ena, Vis parameters of the Entry, depending of wRmCtg
'   (Vis setings not well written, but functionally Ok.)
' FraRmDesc is usd to accept wRmDesc, wRmIntQly. (extention of Grid, recalc)
'   wRmDesc & wRmIntQlt are TxtDesc fields, field_when disables the rest of form
'   CmdDescOk enables the rest of form.
'   The reason for putting these fields below the grid is lack of space in the grid.
'   As far as possible avoid scrolling of the grid.
' X-'', A-RmMst(GrdRm)
'   RowWhen gives value to wRmDesc,..
'
'---
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ms_wRmCtg As String, ms_wRmSrt As String  ' remembering for next entry
Dim wFlag As Boolean, ms_QWModYN As Boolean

Private Sub ADC_Load()
  '*** Set the working field wRmCtg as the First Key Control
  '*** Set the First Non Key Control property to the grid GrdRm
  '*** Set the Child Property Of ADC to GrdRm (it means GrdRm is a child of ADC)
  '*** Set the Previous and Next Control Properties for the Grid GrdRm
      'The Previous Control Property(PrevCtl) is set so that when the user tabs
      'out of the previous control then the focus goes to the Grid.
      'The Next Control Property(NextCtl) is set so that when the user is in the
      'grid and presses Ctrl+W then the focus should go to the Next Control
  '*** Set the Hot Key for the Grid GrdRm as 'A'
      'This is to directly send the focus to the Grid when the user clicks on Alt+A
  '*** If the menu option selected is for Raw Material Master, then
      'make the field RmZCd of the grid as normal
  '*** If the menu option selected is for Customs Code, then
      'make the field RmZCd of the grid as invisible
  
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
  
  Set ADC.FirKeyCtl = ADC("wRmCtg")
  Set ADC.FirNKeyCtl = GrdRm
  ADC.Child = "GrdRm"
  GrdRm.Child = "GrdRmAlyCmp"
  Set GrdRm.PrevCtl = ADC
  Set GrdRm.NextCtl = ADC
  Set GrdRmAlyCmp.PrevCtl = GrdRm
  Set GrdRmAlyCmp.NextCtl = GrdRm
  GrdRm.HotKey = "A"
  GrdRmAlyCmp.HotKey = "B"
  
  '***** Seting RmZcd fgcInvisible for Customs Entry
  '****** Sachin 3.02.0 [26-11-2007] - Purge Partition (Visibility of the 'Push To Current' Button)
  If UCase(ADC.MenuCd) = UCase("RmMst") Then
    GrdRm.ColProp("RmZCd").Style = fgcNormal
    CmdPushCurr.Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
    GrdRm.ColProp("RmZCd").Style = fgcinvisible
    CmdPushCurr.Visible = False
  End If
  
  ms_wRmCtg = ""
  ms_wRmSrt = ""
  
  '*** Jay 3.01 ***
  ' ***** Manali 3.03 - RmId Column property changed to Normal in Grid so following code not reqd
  'If moCn.RecSeek("Select 'x' from SysDb where SysId = 'CRM'") Then GrdRm.ColProp("RmId").Style = fgcNormal
  '*** Jay 3.01 ***
  '****** Sachin 2.13.0 - 05-06-2006 ******
  ms_QWModYN = IIF(moCn.GetFldVal("Select HQWModYN From Head Where HCd= '" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  '****** Sachin 2.13.0 - 05-06-2006 ******
  
  '****** Sachin 3.02.0
  ADC("WRMPRTKEY") = ctCurrPrtn '****** Setting Default Filter to 'Current'
  LblPrtnKey.Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  ADC("WRMPRTKEY").Visible = IIF(gs_Partition = ctCurrPrtn, False, True)
  GrdRm.ColProp("RmPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
  
End Sub

Private Sub ADC_KeyWhen()
  '*** setting the default values for RmCtg and Sort Option from the previous entry
  If ms_wRmCtg <> "" Then ADC("wRmCtg") = ms_wRmCtg
  If ms_wRmSrt <> "" Then ADC("wRmSrt") = ms_wRmSrt
  wFlag = False
  CmdPushCurr.Enabled = False       '****** Sachin 3.02 [26/11/07] - Purge Partition
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the RmCtg and the Sort Option is valid
      '(Sort option can be blank in which case the default sorting is on RmSCtg, RmCd)
      'If not valid then the focus is not allowed to leave the Key Frame
  Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'RMCTG' " + _
                            "And PMCd = '" + ADC("wRmCtg") + "'")
  If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
  
  '***** wRmSrt '' or Valid 'SrtRm' permitted
  If Trim(ADC("wRmSrt")) <> "" Then
    Cancel = Not moCn.RecSeek("Select PMCd From Param Where PTyp = 'SRTRM' " + _
                              "And PMCd = '" + ADC("wRmSrt") + "'")
  Else
    Cancel = False
  End If
  If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(Cancel As Boolean, ErrMsg As String)
  '*** Store the RmCtg, Sort Option for the next entry
  '*** Set the Style Property for the columns in the grid based on the Rm Category
      'fgcNormal= Makes the column Visible & Editable
      'fgcinvisible= Makes the column Invisible & Non-Editable
      'fgcReadOnly= Makes the column Read Only

  ms_wRmCtg = ADC("wRmCtg")
  ms_wRmSrt = ADC("wRmSrt")
  CmdPushCurr.Enabled = True    '****** Sachin 3.02 [26/11/07] - Purge Partition
  
  Select Case ADC("wRmCtg")
  Case Is = "D", "C"
    GrdRm.ColProp("RmSCtg").Style = fgcNormal
    GrdRm.ColProp("RmCd").Style = fgcNormal
    GrdRm.ColProp("RmUom").Style = fgcNormal
    GrdRm.ColProp("RmCol").Style = fgcNormal
    GrdRm.ColProp("RmCmCd").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmZCd").Style = fgcNormal
    GrdRm.ColProp("RmKt").Style = fgcinvisible
    GrdRm.ColProp("RmSpGrav").Style = fgcinvisible
    GrdRm.ColProp("RmPurityRt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityZ").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmFndWt").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmSTWGrp").Style = fgcNormal
    
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    If ms_QWModYN = True Then
      If UCase(ADC.MenuCd) = UCase("RmMst") Then
        GrdRm.ColProp("RmQw").Style = fgcNormal: GrdRm.ColProp("RmQwCst").Style = fgcNormal
      Else
        GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      End If
    Else
      If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
    End If
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    
    GrdRm.ColProp("RmBaseCd").Style = fgcinvisible
    GrdRm.ColProp("RmBasePurityWt").Style = fgcinvisible
    If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmParentCd").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("RmZMst") Then GrdRm.ColProp("RmParentCd").Style = fgcinvisible
    GrdRm.ColProp("RmAlyCd").Style = fgcinvisible     ' ****** Manali 3.8.0
    GrdRm.Height = 7500   '   7590
    GrdRmAlyCmp.Visible = False
    GrdRmAlyCmp.Enabled = False
    ADC("wRaPerGmWt").Visible = False
    'JewelSmith.2 style property for stone class
    GrdRm.ColProp("RmStnCls").Style = fgcNormal
  Case Is = "X"
    GrdRm.ColProp("RmSCtg").Style = fgcNormal
    GrdRm.ColProp("RmCd").Style = fgcNormal
    GrdRm.ColProp("RmUom").Style = fgcNormal
    GrdRm.ColProp("RmKt").Style = fgcNormal
    GrdRm.ColProp("RmCol").Style = fgcNormal
    GrdRm.ColProp("RmCmCd").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmZCd").Style = fgcNormal
    GrdRm.ColProp("RmSpGrav").Style = fgcinvisible
    GrdRm.ColProp("RmPurityRt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityZ").Style = fgcinvisible
    GrdRm.ColProp("RmSTWGrp").Style = fgcNormal   '6.1 new
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    If ms_QWModYN = True Then
      If UCase(ADC.MenuCd) = UCase("RmMst") Then
        GrdRm.ColProp("RmQw").Style = fgcNormal: GrdRm.ColProp("RmQwCst").Style = fgcNormal
      Else
        GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      End If
    Else
      If UCase(ADC.MenuCd) = UCase("RmMst") Then
        GrdRm.ColProp("RmQw").Style = fgcNormal: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      Else
        GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      End If
    End If
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    GrdRm.ColProp("RmBaseCd").Style = fgcNormal
    GrdRm.ColProp("RmBasePurityWt").Style = fgcNormal
    GrdRm.ColProp("RmParentCd").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      GrdRm.ColProp("RmFndWt").Style = fgcNormal
    Else
      GrdRm.ColProp("RmFndWt").Style = fgcinvisible
    End If
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmAlyCd").Style = fgcinvisible     ' ****** Manali 3.8.0
    GrdRm.Height = 7500   '   7590
    GrdRmAlyCmp.Visible = False
    GrdRmAlyCmp.Enabled = False
    ADC("wRaPerGmWt").Visible = False
    'JewelSmith.2 style property for stone class
    GrdRm.ColProp("RmStnCls").Style = fgcinvisible
  Case Is = "M"
    GrdRm.ColProp("RmSCtg").Style = fgcNormal
    GrdRm.ColProp("RmCd").Style = fgcNormal
    GrdRm.ColProp("RmUom").Style = fgcNormal
    GrdRm.ColProp("RmKt").Style = fgcNormal 'fgcinvisible
    
    GrdRm.ColProp("RmCol").Style = fgcNormal
    GrdRm.ColProp("RmCmCd").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmZCd").Style = fgcNormal
    GrdRm.ColProp("RmSpGrav").Style = fgcNormal
    GrdRm.ColProp("RmPurityRt").Style = fgcNormal
    GrdRm.ColProp("RmPurityWt").Style = fgcNormal
    GrdRm.ColProp("RmPurityZ").Style = fgcNormal
    GrdRm.ColProp("RmSTWGrp").Style = fgcNormal '6.1 new
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    If ms_QWModYN = True Then
      If UCase(ADC.MenuCd) = UCase("RmMst") Then
        GrdRm.ColProp("RmQw").Style = fgcNormal: GrdRm.ColProp("RmQwCst").Style = fgcNormal
      Else
        GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      End If
    Else
      If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
    End If
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    GrdRm.ColProp("RmBaseCd").Style = fgcinvisible
    GrdRm.ColProp("RmBasePurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmParentCd").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmFndWt").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmAlyCd").Style = fgcinvisible     ' ****** Manali 3.8.0
    GrdRm.Height = 7500   '   7590
    GrdRmAlyCmp.Visible = False
    GrdRmAlyCmp.Enabled = False
    ADC("wRaPerGmWt").Visible = False
    'JewelSmith.2 style property for stone class
    GrdRm.ColProp("RmStnCls").Style = fgcinvisible
  Case Is = "A"
    GrdRm.ColProp("RmSCtg").Style = fgcNormal
    GrdRm.ColProp("RmCd").Style = fgcNormal
    GrdRm.ColProp("RmUom").Style = fgcNormal
    GrdRm.ColProp("RmKt").Style = fgcinvisible
    GrdRm.ColProp("RmCol").Style = fgcNormal
    GrdRm.ColProp("RmCmCd").Style = fgcNormal
    GrdRm.ColProp("RmSpGrav").Style = fgcNormal
    GrdRm.ColProp("RmPurityRt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmPurityZ").Style = fgcinvisible
    GrdRm.ColProp("RmZCd").Style = fgcinvisible
    GrdRm.ColProp("RmSTWGrp").Style = fgcinvisible
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    GrdRm.ColProp("RmBaseCd").Style = fgcinvisible
    GrdRm.ColProp("RmBasePurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmParentCd").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.ColProp("RmFndWt").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    GrdRm.Height = 4950
    GrdRmAlyCmp.Visible = True
    GrdRmAlyCmp.Enabled = True
    ADC("wRaPerGmWt").Visible = True
  ' ******* Manali - 3.03 - paLadium added
    'JewelSmith.2 style property for stone class
    GrdRm.ColProp("RmStnCls").Style = fgcinvisible
   Case Is = "G", "P", "S", "L"
    GrdRm.ColProp("RmSCtg").Style = fgcNormal
    GrdRm.ColProp("RmCd").Style = fgcNormal
    GrdRm.ColProp("RmUom").Style = fgcNormal
    GrdRm.ColProp("RmKt").Style = fgcNormal
    GrdRm.ColProp("RmCol").Style = fgcNormal
    GrdRm.ColProp("RmCmCd").Style = fgcNormal
    If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmZCd").Style = fgcNormal
    GrdRm.ColProp("RmPurityRt").Style = fgcNormal
    GrdRm.ColProp("RmPurityWt").Style = fgcNormal
    GrdRm.ColProp("RmPurityZ").Style = fgcNormal
    GrdRm.ColProp("RmSpGrav").Style = fgcNormal
    GrdRm.ColProp("RmSTWGrp").Style = fgcNormal '6.1 new
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    If ms_QWModYN = True Then
      If UCase(ADC.MenuCd) = UCase("RmMst") Then
        GrdRm.ColProp("RmQw").Style = fgcNormal: GrdRm.ColProp("RmQwCst").Style = fgcNormal
      Else
        GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
      End If
    Else
      If UCase(ADC.MenuCd) = UCase("RmMst") Then GrdRm.ColProp("RmQw").Style = fgcinvisible: GrdRm.ColProp("RmQwCst").Style = fgcinvisible
    End If
    ' ****** Sachin 2.13.0 - 29/05/06 ******
    GrdRm.ColProp("RmBaseCd").Style = fgcinvisible
    GrdRm.ColProp("RmBasePurityWt").Style = fgcinvisible
    GrdRm.ColProp("RmParentCd").Style = fgcinvisible
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      GrdRm.ColProp("RmFndWt").Style = fgcNormal
      GrdRm.ColProp("RmFndWt").Heading = "Per Pc Wt"
      GrdRm.ColProp("RmAlyCd").Style = fgcNormal     ' ****** Manali 3.8.0
    Else
      GrdRm.ColProp("RmFndWt").Style = fgcinvisible
      GrdRm.ColProp("RmAlyCd").Style = fgcinvisible     ' ****** Manali 3.8.0
    End If
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    ' Bef 3.8.0
    ''GrdRm.ColProp("RmFndWt").Style = fgcinvisible
    ' ****** Sachin 2.12 - 06/12/05 - Fnd Wt/Pc
    
    GrdRm.Height = 7500   '   7590
    GrdRmAlyCmp.Visible = False
    GrdRmAlyCmp.Enabled = False
    ADC("wRaPerGmWt").Visible = False
    'JewelSmith.2 style property for stone class
    GrdRm.ColProp("RmStnCls").Style = fgcinvisible
  End Select
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Do not allow the user to save the entry if sum(PerGmWt) for an alloy > 1
      
  If wFlag = True Then pr_Cancel = True: pr_ErrMsg = "Total Per Gram Weight In The Alloy Composition Grid Should Not Be > 1"
End Sub

Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Set Helps For Each Field
  '*** For the help of working field wRmCdFr ('RmCode From' in the selection range option)
      'give the first parameter as 'N' (i.e. select RmCodes which are not custom codes) and
      'the second parameter as the RmCtg
  '*** Similarly for the help of working field wRmCdTo ('RmCode To' in the selection range option)
  
  Select Case IdName
  Case Is = UCase("wRmCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("wRmSCtgFr")
    Call HlpList.PSCd("RMSCTG", ADC("wRmCtg"))
  Case Is = UCase("wRmSCtgTo")
    Call HlpList.PSCd("RMSCTG", ADC("wRmCtg"))
  '*** Jay 2.14 ***(Valid Help Set)
  Case Is = UCase("wRmCdFr")
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      Call HlpList.RmCd("N", "'" + ADC("wRmCtg") + "'", , False)
    ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
      Call HlpList.RmCd("Y", "'" + ADC("wRmCtg") + "'", , False)
    End If
  Case Is = UCase("wRmCdTo")
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      Call HlpList.RmCd("N", "'" + ADC("wRmCtg") + "'", , False)
    ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
      Call HlpList.RmCd("Y", "'" + ADC("wRmCtg") + "'", , False)
    End If
    
  '4.1.3.0
  Case Is = UCase("wCmCdFr"), UCase("wCmCdTo"), UCase("WCMCDSEL")
    HlpList.MultiSelect = IIF(UCase(IdName) = UCase("WCMCDSEL"), True, False)
    Call HlpList.CustCd("C")
        
  '*** Jay 2.14 ***(Valid Help Set)
  Case Is = UCase("wRmSrt")
    Call HlpList.PMCd("SRTRM")
    
  Case Is = UCase("WCUSTYN")            '4.1.3.0
    Call HlpList.PMCd("YN")
    
   
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Dim i As Integer, wTotWtPerGm As Single
  Select Case UCase(IdName)
  Case Is = UCase("wRaPerGmWt")
    If ADC.Mode = xNorm And (GrdRmAlyCmp.Mode = fgmnorm Or GrdRmAlyCmp.Mode = fgmdel Or GrdRmAlyCmp.Mode = fgminit) Then
      wTotWtPerGm = 0
      For i = 1 To GrdRmAlyCmp.Rows - 1
        If Not GrdRmAlyCmp.IsDel(i) Then
          wTotWtPerGm = wTotWtPerGm + GrdRmAlyCmp.Value(i, "RaPerGmWt")
        End If
      Next i
      ADC("wRaPerGmWt") = wTotWtPerGm
    End If
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wRmCtg       valid PMCd('RMCTG')
  '*** wRmSrt       valid PMCd('SRTRM') or ''
  
  Select Case UCase(IdName)
  Case Is = UCase("wRmCtg")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' " + _
                         "and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Raw Material Category": Exit Sub
    
  '4.1.3.0
  Case Is = UCase("wCustYn")
    Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' " + _
                         "and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Customer YN Option": Exit Sub
    
  Case Is = UCase("wRmSrt")
    Cancel = Trim(pv_NewValue) <> "" And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                         "PTyp= 'SRTRM' and PMCd = '" + pv_NewValue + "'"))
    ErrMsg = "Invalid Sort Option": Exit Sub
  End Select
End Sub

Private Sub CmdPushCurr_Click()
    '****** Sachin 3.02.0 [26-11-2007] - Purge Partition
    Call PushToCurrent

End Sub

Private Sub GrdRm_RowWhen(ByVal RowNum As Integer)
  '*** give the default value for the RmZ field based on the menu code
      'give 'Y' or 'N' depending on whether the code is a Raw Material Code or Customs Code
  '*** Set values for the working fields in the description frame

  With GrdRm
  
    '****** Sachin 3.02 18-12-2007 - Setting the Default Value to Current ******
    If .IsNew(RowNum) Then .Value(RowNum, "RmPrtKey") = ctCurrPrtn
    '****** Sachin 3.02 18-12-2007 ******
  
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      .Value(RowNum, "RmZ") = "N"
      CmdPushCurr.Enabled = IIF(.Value(RowNum, "RmPrtKey") <> ctCurrPrtn, True, False)
    ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
      .Value(RowNum, "RmZ") = "Y"
    End If
        
    ADC("wRmDesc") = GrdRm.Value(RowNum, "RmDesc")
    ADC("wRmIntQly") = GrdRm.Value(RowNum, "RmIntQly")
    ADC("wRmPrtDesc") = GrdRm.Value(RowNum, "RmPrtDesc")
    ' ###################################  Manoj #### Ver: 2.0.6  ###################################
    ' Date:   20/01/2004
    If .IsNew(RowNum) Then .Value(RowNum, "RmValidYN") = "Y"
    ADC("wRmValidYN") = .Value(RowNum, "RmValidYN")
    ' ###########################################  Manoj  ###########################################
    If GrdRmAlyCmp.Rows - 1 = 0 Then ADC("wRaPerGmWt") = 0
    
    '****** Sachin 2.13.0 - 20-06-2006 ******
    If .IsNew(RowNum) Then .Value(RowNum, "RmQw") = "W": .Value(RowNum, "RmQwCst") = "W"
    '****** Sachin 2.13.0 - 20-06-2006 ******
    'JewelSmith.2 default stone class value for D and C
    If .IsNew(RowNum) Then
      If ADC("wRmCtg") = "C" Or ADC("wRmCtg") = "D" Then .Value(RowNum, "RmStnCls") = "-"
    End If
    
  End With
  
End Sub
Private Sub GrdRm_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** The help for RmCd is based on the menu code
      'If the entry is for custom codes then pass 'Y' as the first parameter else pass 'N'
      'pass the RmCtg as the next parameter
  '*** For the help of RmZCd, pass 'Y' as the first parameter and the RmCtg as the second parameter
  '*** Set the values for RmSpGrav, RmPurityRt, RmPurityWt as 1 if the current value is 0
  
  '****** Sachin 3.02 [30/11/07] - Purge Partition - Locking Previous Partition Records
  If GrdRm.Value(RowNum, "RmPrtKey") <> ctCurrPrtn Then
    Cancel = True: ErrMsg = "Cannot Edit Records From Previous Partition"
    FraRmDesc.Enabled = False
    Exit Sub
  Else
    FraRmDesc.Enabled = True
  End If
  '****** Sachin 3.02 [30/11/07] - Purge Partition - Locking Previous Partition Records
  
  Select Case ColName
  Case Is = UCase("RmSCtg")
    Call HlpList.PSCd("RMSCTG", ADC("wRmCtg"))
  Case Is = UCase("RmCd")
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      Call HlpList.RmCd("N", "'" + ADC("wRmCtg") + "'", GrdRm.Value(RowNum, "RmSCtg"), False)
    ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
      Call HlpList.RmCd("Y", "'" + ADC("wRmCtg") + "'", GrdRm.Value(RowNum, "RmSCtg"), False)
    End If
  Case Is = UCase("RmCmCd")
    Call HlpList.CustCd("C")
  Case Is = UCase("RmUom")
    Call HlpList.PMCd("RMUOM")
  Case Is = UCase("RmCol")
    Call HlpList.PSCd("RMCOL", ADC("wRmCtg"))
  Case Is = UCase("RmKt")
    Call HlpList.PMCd("KT")
  Case Is = UCase("RmZCd")
    Call HlpList.RmCd("Y", "'" + ADC("wRmCtg") + "'")
  Case Is = UCase("RmSpGrav")
    If GrdRm.Value(RowNum, "RmSpGrav") = 0 Then _
       GrdRm.Value(RowNum, "RmSpGrav") = 1
  Case Is = UCase("RmPurityRt")
    If GrdRm.Value(RowNum, "RmPurityRt") = 0 Then _
       GrdRm.Value(RowNum, "RmPurityRt") = 1
  Case Is = UCase("RmPurityWt")
' **** Zubin 211 **** ' (Remove below when not required)
'    If moCn.RecSeek("Select 'x' From Txnd Where TdRmCd= '" + GrdRm.Value(RowNum, "RmCd") + "'") Then
'      Cancel = True: ErrMsg = "Cannot Edit Production Purity Since Transaction Exists For This RmCd": Exit Sub
'    End If
' **** Zubin 211 **** '
    If GrdRm.Value(RowNum, "RmPurityWt") = 0 Then _
       GrdRm.Value(RowNum, "RmPurityWt") = 1
'' ****** Sachin 2.12 - 05/12/05 - SJM
    If (ADC("wRmCtg") = "S") And ((moCn.GetFldVal("Select HSlvModYN From Head Where HCoCd = '" + gs_CoCd + "'") = "N")) Then
      Cancel = True: ErrMsg = "Cannot Change Prod. Purity for Silver as Silver Module Option is set to NO in Head": Exit Sub
    End If
'' ****** Sachin 2.12 - 05/12/05 - SJM
' ******* Manali - 3.03 - 16/06/08 - Paladium Module
    If (ADC("wRmCtg") = "L") And ((moCn.GetFldVal("Select HPdModYN From Head Where HCoCd = '" + gs_CoCd + "'") = "N")) Then
      Cancel = True: ErrMsg = "Cannot Change Prod. Purity for Paladium as paLadium Module Option is set to NO in Head": Exit Sub
    End If
' ******* Manali - 3.03 - 16/06/08 - Paladium Module
  Case Is = UCase("RmPurityZ")
    If GrdRm.Value(RowNum, "RmPurityZ") = 0 Then _
       GrdRm.Value(RowNum, "RmPurityZ") = 1
  Case Is = UCase("wDesc")
    ADC("wRmDesc").SetFocus
  Case Is = UCase("RmSTWGrp")
    Call HlpList.PMCd("STWGRP")
  Case Is = UCase("RmQw"), UCase("RmQwCst")
    Call HlpList.PMCd("LABQW")
  Case Is = UCase("RmBaseCd")
    ' Sachin 212
    '***** Manali 3.03 - 16/06/08 - paLadium included
    Call HlpList.RmCd("N", "'G','P','S', 'L'")
  Case Is = UCase("RmParentCd")
    ' **** Manali 3.6.0 - 29/01/10
    Call HlpList.PMCd("PARENTCD", "'" + ADC("wRmCtg") + "'")
    '''Call HlpList.PMCd("PARENTCD", "'" + ADC("wRmCtg") + "'", GrdRm.Value(RowNum, "RmSCtg"))
  Case Is = UCase("RmId")
    Call HlpList.PSCd("RMID", ADC("wRmCtg"))
    
  ' ***** Manali 3.8.0
  Case Is = UCase("RmAlyCd")
    Call HlpList.RmCd("N", "'A'")
  'JewelSmith.2 calling help of stone class
  Case Is = UCase("RmStnCls")
    Call HlpList.PSCd("STNCLS", ADC("wRmCtg"))
    
  End Select
End Sub

Private Sub GrdRm_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** RmBasePurityWt on RmBaseCd               Normal
      'RmBasePurityWt= RmPurityWt from RmMst(RmCd= RmBaseCd of Grid)
  '*** RmQwCst on RmQw               Normal
      'RmQwCst= RmQw If RmQwCst is Blank
  
  With GrdRm
    Select Case UCase(ColName)
    Case Is = UCase("RmBasePurityWt")
      '****** Sachin 3.02.0 [30/11/07] - Purge Partition (Partition Key added)
      If .Mode = fgmnorm Then
        .Value(RowNum, "RmBasePurityWt") = moCn.GetFldVal("Select RmPurityWt from RmMst where " + _
                                           "RmCd= '" + .Value(RowNum, "RmBaseCd") + "' And RmPrtkey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' Then '" + ctCurrPrtn + "' Else RmPrtKey End)")
      End If
      
    '****** Sachin 2.13.0 - 29-05-2006 ******
    Case Is = UCase("RmQwCst")
      If .Mode = fgmnorm Then
        If (.Value(RowNum, "RmQwCst") = "" Or (ms_QWModYN = False And ADC("wRmCtg") = "X")) Then .Value(RowNum, "RmQwCst") = .Value(RowNum, "RmQw")
      End If
    '****** Sachin 2.13.0 - 29-05-2006 ******
    
    End Select
  End With
End Sub

Private Sub GrdRm_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** RmSCtg         valid PSCd('RMSCTG', RmCtg)
  '*** RmCd           Not Blank
  '*** RmUOM          valid PMCd('RMUOM')
  '*** RmCol          valid PSCd('RMCOL', RmCtg)
  '*** RmKT           valid PMCd('KT')
  '*** RmSpGrav       <> 0
  '*** RmPurityRt     <> 0
  '*** RmPurityWt     <> 0
  '*** RmZCd          valid RmCd from RmMst(RmCtg= adc("wRmCtg") and RmZ= 'Y')
  
  With GrdRm
    Select Case ColName
    Case Is = UCase("RmSCtg")
      Cancel = Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                           "and PMCd = '" + ADC("wRmCtg") + "' and " + _
                           "PSCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
    Case Is = UCase("RmCd")
      If Trim(NewValue) = "" Then _
      Cancel = True: ErrMsg = "Invalid Raw Material Code": Exit Sub
    Case Is = UCase("RmCmCd")
      If Trim(NewValue) <> "" Then
        Cancel = Not moCn.RecSeek("Select CmCd From CustMst where CmCtg= 'C' " + _
                             "and CmCd = '" + NewValue + "'")
        ErrMsg = "Invalid Customer Code": Exit Sub
      End If
    Case Is = UCase("RmUom")
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMUOM' " + _
                           "and PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Unit Of Measurement": Exit Sub
    Case Is = UCase("RmCol")
      Cancel = Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMCOL' " + _
                           "and PMCd = '" + ADC("wRmCtg") + "' and " + _
                           "PSCd = '" + NewValue + "'")
      ErrMsg = "Invalid Raw Material Colour": Exit Sub
    Case Is = UCase("RmKt")
      If Trim(NewValue) <> "" Then
        Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'KT' " + _
                             "and PMCd = '" + NewValue + "'")
        ErrMsg = "Invalid Raw Material Karat": Exit Sub
      End If
    Case Is = UCase("RmSpGrav")
      If NewValue = 0 Then
        Cancel = True: ErrMsg = "Specific Gravity Cannot Be Zero": Exit Sub
      End If
    Case Is = UCase("RmPurityRt")
      If NewValue = 0 Then
        Cancel = True: ErrMsg = "Purity Cannot Be Zero": Exit Sub
      End If
    Case Is = UCase("RmPurityWt")
      If NewValue = 0 Then
        'Cancel = True: ErrMsg = "Production Purity Cannot Be Zero": Exit Sub
      End If
    Case Is = UCase("RmPurityZ")
      If NewValue = 0 Then
        Cancel = True: ErrMsg = "Customs Purity Cannot Be Zero": Exit Sub
      End If
    Case Is = UCase("RmZCd")
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
      '****** Sachin 3.02.0 - [30/11/07] - Purge Partition (Partition Key added)
      Cancel = Not moCn.RecSeek("Select RmCd From RmMst where " + _
                           "RmZ='Y' and RmCtg= '" + ADC("wRmCtg") + "' " + _
                           "and RmCd = '" + NewValue + "' And RmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else RmPrtKey End)")
      ErrMsg = "Invalid Customs Code": Exit Sub
'      Cancel = Not moCn.RecSeek("Select RmCd From RmMst where " + _
'                           "RmZ='Y' and RmCtg= '" + ADC("wRmCtg") + "' " + _
'                           "and RmCd = '" + NewValue + "'")
'      ErrMsg = "Invalid Customs Code": Exit Sub
' ###########################################  Manoj  ###########################################

    Case Is = UCase("RmSTWGrp")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where " + _
                                    "PTyp= 'STWGRP' and PmCd = '" + NewValue + "'"))
      ErrMsg = "Invalid Size To Weight Group": Exit Sub
      
    '****** Sachin 2.13.0 - 29-05-2006 ******
    Case Is = UCase("RmQw")
      Cancel = Not moCn.RecSeek("Select PmCd From Param where " + _
                                    "PTyp= 'LABQW' and PmCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Rm Sale Q/W": Exit Sub
    
    Case Is = UCase("RmQwCst")
      Cancel = Not moCn.RecSeek("Select PmCd From Param where " + _
                                    "PTyp= 'LABQW' and PmCd = '" + NewValue + "'")
      If Cancel = True Then ErrMsg = "Invalid Rm Cost Q/W": Exit Sub
    '****** Sachin 2.13.0 - 29-05-2006 ******
    
    Case Is = UCase("RmBaseCd")
    '****** Sachin 3.02 [26/11/07] - Speed Optimization - Purge Partition
    ' ***** Manali 3.03 - 16/06/08 - paLadium included
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst where " + _
                                    "RmZ='N' and RmCtg in ('G', 'P', 'S', 'L') And RmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else RmPrtKey End)" + _
                                    "and RmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Base Code (Base Code Should Be Of Category 'G', 'P', 'S' or 'L')": Exit Sub
      
    Case Is = UCase("RmParentCd")
      ' **** Manali 3.6.0 - 29/01/10
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where " + _
                                    "PTyp= 'PARENTCD' and PmCd = '" + NewValue + "' and " + _
                                    "PValue= '" + ADC("wRmCtg") + "' "))
''      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where " + _
''                                    "PTyp= 'PARENTCD' and PmCd = '" + NewValue + "' and " + _
''                                    "PValue= '" + ADC("wRmCtg") + "' and PValue1= '" + GrdRm.Value(RowNum, "RmSCtg") + "' "))
      If Cancel = True Then ErrMsg = "Invalid Parent Code": Exit Sub
      
    Case Is = UCase("RmId")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From Param where " + _
                                    "PTyp= 'RMID' and PMCd= '" + ADC("wRmCtg") + "' and " + _
                                    "PSCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid RM Id": Exit Sub
    
    Case Is = UCase("RmAlyCd")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select 'x' From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + _
                                    "RmCtg= 'A' and RmCd= '" + NewValue + "' and " + _
                                    "RmZ='N' and RmValidYn='Y' "))
      If Cancel = True Then ErrMsg = "Invalid Alloy Code": Exit Sub
    'JewelSmith.2 validation for stone class
    Case Is = UCase("RmStnCls")
      Cancel = Not moCn.RecSeek("Select 'x' From Param where " + _
                                    "PTyp= 'STNCLS' and PMCd= '" + ADC("wRmCtg") + "' and " + _
                                    "PSCd = '" + NewValue + "' ")
      If Cancel = True Then ErrMsg = "Invalid Stone Class": Exit Sub
            
    End Select
  End With
End Sub
Private Sub GrdRm_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  With GrdRm
    .Store "RmCtg", ADC("wRmCtg")
    If .IsNew(RowNum) Then .Store "RmDt", Format$(moCn.SrvrDate, "DD-MM-YY")   '****** Sachin 3.02 - Raw Material Created Date
  End With
End Sub
Private Sub GrdRmAlyCmp_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  '*** The help for Composite RmCd is based on the menu code
      'If the entry is for custom codes then pass 'Y' as the first parameter else pass 'N'
      'pass the RmCtg as the next parameter
  '*** For the help of RmZCd, pass 'Y' as the first parameter and the RmCtg as the second parameter
  '*** Set the values for RmSpGrav, RmPurityRt, RmPurityWt as 1 if the current value is 0
  
  With GrdRmAlyCmp
  Select Case ColName
  Case Is = UCase("RaCmpCtg")
    Call HlpList.PMCd("RMCTG")
  Case Is = UCase("RaCmpSCtg")
    Call HlpList.PSCd("RMSCTG", .Value(RowNum, "RaCmpCtg"))
  Case Is = UCase("RaCmpCd")
    If UCase(ADC.MenuCd) = UCase("RmMst") Then
      Call HlpList.RmCd("N", "'" + .Value(RowNum, "RaCmpCtg") + "'", .Value(RowNum, "RaCmpSCtg"))
    ElseIf UCase(ADC.MenuCd) = UCase("RmZMst") Then
      Call HlpList.RmCd("Y", "'" + .Value(RowNum, "RaCmpCtg") + "'", .Value(RowNum, "RaCmpSCtg"))
    End If
  End Select
  End With
End Sub
Private Sub GrdRmAlyCmp_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** RaCmpCtg          valid PMCd('RMCTG')
  '*** RaCmpSCtg         valid PSCd('RMSCTG', RaCmpCtg)
  '*** RaCmpCd           valid RmCd From RmMst(RaCmpCtg, RaCmpSCtg)
  
  With GrdRmAlyCmp
    Select Case ColName
    Case Is = UCase("RaCmpCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'RMCTG' " + _
                                "and PMCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Category For Alloy Composite Code": Exit Sub
    Case Is = UCase("RaCmpSCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'RMSCTG' " + _
                                "and PMCd = '" + .Value(RowNum, "RaCmpCtg") + "' and " + _
                                "PSCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Sub Category For Alloy Composite Code": Exit Sub
    Case Is = UCase("RaCmpCd")
      '****** Sachin 3.02 [26/11/07] - Speed Optimization - Purge Partition
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select RmCd From RmMst Where RmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else RmPrtKey End)" + _
                                    " And RmCtg= '" + .Value(RowNum, "RaCmpCtg") + "' and " + _
                                    "RmSCtg= '" + .Value(RowNum, "RaCmpSCtg") + "' and " + _
                                    "RmZ = '" + IIF(UCase(ADC.MenuCd) = UCase("RmMst"), "N", "Y") + "'"))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Alloy Composite Code": Exit Sub
    Case Is = UCase("RaPerGmWt")
      Cancel = Round(NewValue, 3) < 0 Or Round(NewValue, 3) > 1
      If Cancel = True Then ErrMsg = "Per Gram Weight Should Be <= 1": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdRmAlyCmp_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
      'The fields of the Base table which are not existing in the grid are stored in the
      'table using Grid_InitKey
  With GrdRmAlyCmp
    .Store "RaCd", GrdRm.Value(GrdRm.Row, "RmCd")
  End With
End Sub
Private Sub GrdRmAlyCmp_LostFocus()
  wFlag = False
  Dim wTotRmWt As Single
  If GrdRmAlyCmp.Rows - 1 > 0 Then
    wTotRmWt = moCn.GetFldVal("Select sum(RaPerGmWt) from RmAlyCmp where " + _
                           "RaCd= '" + GrdRm.Value(GrdRm.Row, "RmCd") + "'")
    If Round(wTotRmWt, 3) < 0 Or Round(wTotRmWt, 3) > 1 Then
      wFlag = True
      DispMsg "Total Of Per Gram Weight In The Alloy Composition Grid Should Be >= 0 And <= 1 ", etError
      GrdRmAlyCmp.SetFocus: Exit Sub
    End If
  End If
End Sub
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
Private Sub TXTDESC_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("wRmValidYN")
    Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' " + _
                         "and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Valid (Y/N)": Exit Sub
  End Select
End Sub
' ###########################################  Manoj  ###########################################

Private Sub TXTDESC_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** When any of the working fields of FraRmDesc frame gets the focus then disable
      'all the other frames and also the ADC
  '*** Set Help For the working field of International Quality
  ADC.AllowSave = False
  FraNKeyAll.Enabled = False
  Select Case IdName
  Case Is = UCase("wRmIntQly")
    Call HlpList.RmIntQly
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
  Case Is = UCase("wRmValidYN")
    Call HlpList.PMCd("YN")
' ###########################################  Manoj  ###########################################
  End Select
End Sub
Private Sub TXTDESC_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  
  'ADC.AllowSave = False
  'FraNKeyAll.Enabled = False
  Select Case IdName
  Case Is = UCase("wRmPrtDesc")
    If ADC.Mode = xNorm Then
      If ADC("wRmPrtDesc") = "" Then ADC("wRmPrtDesc") = ADC("wRmDesc")
    End If
  End Select
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdRm_SetRecSource()
  '*** Set the Record Source of the Grid GrdRm
      'If the user has not specified the sort option then the default sort option
      'is RmSCtg, RmCd
  '*** The select statement is based on the menu code (i.e. Rm Codes or Custom Codes)
  
  Dim ws_Cnd As String, ws_OrdBy As String, ws_RmZ As String
  Dim ws_CmCnd As String        '4.1.3.0
  '***** CmpStr property is set for scope fields in the Key Frame. So whatever is the
        'record selection condition specified by the user can be obtained using ADC.RepCond
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  If ADC("wRmSrt") = "" Then
    ws_OrdBy = " Order By RmSCtg, RmCd "
  Else
    ws_OrdBy = " Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
               " Where PTyp='SRTRM' and PMCd='" + ADC("wRmSrt") + "'")
  End If
  
  ws_RmZ = IIF(UCase(ADC.MenuCd) = UCase("RmMst"), "N", "Y")
  
  '4.1.3.0
  If UCase(ADC.MenuCd) = UCase("RmMst") Then _
     ws_CmCnd = IIF(UCase(ADC("WCUSTYN") = "Y"), " and RmCmCd<>'' ", IIF(UCase(ADC("WCUSTYN")) = "N", " and RmCmCd='' ", " ")) + _
                IIF(UCase(ADC("WCMCDSEL") <> ""), "And '," + ADC("WCMCDSEL") + ",' Like '%,'+RmCmCd+',%' ", "")
  
  '****** Sachin 3.02 [26/11/07] - Speed Optimization (Purge Partition)
  'And RmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else RmPrtKey End)"
  GrdRm.RecSource = " Select * from RmMst where RmZ='" + ws_RmZ + "' " + ws_Cnd + ws_CmCnd + ws_OrdBy
  
End Sub
Private Sub GrdRmAlyCmp_SetRecSource()
  '*** Set the Record Source of the Grid GrdRmAlyCmp
  GrdRmAlyCmp.RecSource = " Select * from RmAlyCmp where RaCd= '" + GrdRm.Value(GrdRm.Row, "RmCd") + "' Order By RaCmpCtg, RaCmpSCtg, RaCmpCd "
End Sub
Private Sub CmdDescOk_Click()
  '*** When the user clicks clicks on this command button, he comes out of the description
      'frame(FraRmDesc) and the values in the working fields of the frame are copied into
      'the columns of the grid (only if the values are different)
  Dim wRow
  With GrdRm
    wRow = IIF(.Row > 0, .Row, 0)
    '*** TXTDESC(0) is working field for RmDesc
    '*** TXTDESC(1) is working field for RmIntQly
    If wRow > 0 Then
      If .Value(wRow, "RmDesc") <> TXTDESC(0).TxtText Then _
         .Value(wRow, "RmDesc") = TXTDESC(0).TxtText
      If .Value(wRow, "RmIntQly") <> TXTDESC(1).TxtText Then _
         .Value(wRow, "RmIntQly") = TXTDESC(1).TxtText
      If .Value(wRow, "RmPrtDesc") <> TXTDESC(2).TxtText Then _
         .Value(wRow, "RmPrtDesc") = TXTDESC(2).TxtText
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
      If .Value(wRow, "RmValidYN") <> TXTDESC(3).TxtText Then _
         .Value(wRow, "RmValidYN") = TXTDESC(3).TxtText
' ###########################################  Manoj  ###########################################
    End If
  End With
  
  FraNKeyAll.Enabled = True
  GrdRm.Col = 1
  GrdRm.SetFocus
  ADC.AllowSave = True
  
  '**** fun *****
  Dim wFunny As MDORowSet
  Set wFunny = moCn.OpenResultset("Select RmCd as qRmCd from RmMst where RmCd = 'NRND' And RmPrtKey = (Case When '" + gs_Partition + "'='" + ctCurrPrtn + "' then '" + ctCurrPrtn + "' Else RmPrtKey End)")
  If wFunny.EOF = False And wFunny.BOF = False Then
    wFunny.edit
    wFunny!qRmCmCd = "A"
    'wFunny.Update
  End If
  '**** fun *****
End Sub

'''  ***** Std Code not to be changed **********************************
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Color
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
  Dim ws_BtnStr As String
  ws_BtnStr = "CmdDescOk"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Color
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
Private Sub TXTDESC_Validate(Index As Integer, Cancel As Boolean)
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    Set moCn = Nothing    '*** (09/08/05)    '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdRm_Validate(Cancel As Boolean)
  Cancel = GrdRm.Validate
End Sub
Private Sub GrdRmAlyCmp_Validate(Cancel As Boolean)
  Cancel = GrdRmAlyCmp.Validate
End Sub

Private Sub PushToCurrent()
    '****** Sachin 3.02.0 [26-11-2007] - Purge Partition (Pushing RM Code From Previous to Current Partition)
    On Error GoTo errhld
    Dim wOldPrtKey As String
    Dim ws_TblNm As String, ws_QryStr As String, wrsRrTcTyp As MDORowSet
  
      'If moCn.RecSeek(" Select 'x' From RmMst Rm " + _
                      " Left Outer Join RmMst B On B.RmCd=Rm.RmBaseCd and Rm.RmBaseCd<>'' " + _
                      " Where Rm.RmZ='N' and Rm.RmCd='" + GrdRm.Value(GrdRm.Row, "RmCd") + "' And Rm.RmValidYN='N'") Then _
      'DispMsg "Cannot Push An InValid RmCode", etError: Exit Sub
  
      'Temporary Table Names
      ws_TblNm = "#Tmp" + ADC.MenuCd
      wOldPrtKey = GrdRm.Value(GrdRm.Row, "RmPrtKey")
  
      With GrdRm
        MWLib.BeginProcess Me, "Pushing Raw Material Code(s) ..."
        moCn.BeginTrans (mdoCommitRec)
        'Creating Temporary Tables
        ws_QryStr = "Select Rm.RmCd as qRmCd,Rm.RmZCd qRmZCd, " + _
                    " Rm.RmBaseCd as qRmBaseCd,isNull(B.RmZCd,'') qRmBaseZCd " + _
                    "   Into " + ws_TblNm + _
                    "   From RmMst Rm " + _
                    " Left Outer Join RmMst B On B.RmCd=Rm.RmBaseCd and Rm.RmBaseCd<>'' " + _
                    " Where Rm.RmZ='N' and Rm.RmCd='" + .Value(.Row, "RmCd") + "'"
      
        moCn.CreateTmpTable ws_TblNm, ws_QryStr
        
        moCn.Execute (" Update RmMst Set RmPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                      " Where RmPrtKey<>'" + ctCurrPrtn + "' and (RmCd in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        Set wrsRrTcTyp = moCn.OpenRes("Select PMCd From Param Where PTyp='RRTCTYP' ")
        If wrsRrTcTyp.RecCount > 0 Then
          wrsRrTcTyp.MoveFirst
          Do While Not (wrsRrTcTyp.BOF Or wrsRrTcTyp.EOF)
            moCn.Execute (" Update RmRt Set ModUsr='" + gs_UsrCd + "', RrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                         " Where RrTctyp='" + wrsRrTcTyp!PMCd + "' And RrPrtKey<>'" + ctCurrPrtn + "' and (RrCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")
                                     
            moCn.Execute (" Update RmRtHist Set ModUsr='" + gs_UsrCd + "', RhPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                        " Where RhTctyp='" + wrsRrTcTyp!PMCd + "' And RhPrtKey<>'" + ctCurrPrtn + "' and (RhCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")
            wrsRrTcTyp.MoveNext
          Loop
        End If

        'moCn.Execute (" Update RmRt Set RrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                     " Where RrPrtKey<>'" + ctCurrPrtn + "' and (RrCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        'moCn.Execute (" Update RmRtHist Set RhPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                     " Where RhPrtKey<>'" + ctCurrPrtn + "' and (RhCd  in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        moCn.Execute (" Update CustRm Set CrPrtKey='" + ctCurrPrtn + "' From " + ws_TblNm + " Rm " + _
                     " Where CrPrtKey<>'" + ctCurrPrtn + "' and (CrOurRmCd in (qRmCd,qRmZCd,qRmBaseCd,qRmBaseZCd))")

        moCn.CommitTrans mdoCommitRec
        GrdRm.Init
        GrdRm.RefreshData
        CmdPushCurr.Enabled = False
        MWLib.EndProcess Me
    End With
    Exit Sub
    
errhld:
    moCn.RollbackTrans mdoCommitNone
    MsgBox Err.Description & Chr(13) & " Could Not Push The RmCode to Current"
    MWLib.EndProcess Me
    '****** Sachin 3.02.0 [26-11-2007] - Purge Partition (Pushing RM Code From Previous to Current Partition)
End Sub
