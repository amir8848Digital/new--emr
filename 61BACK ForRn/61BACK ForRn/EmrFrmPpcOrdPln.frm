VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrmPpcOrdPln 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Master"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   405
      Left            =   1500
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   9480
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_BTN1 CmdAutoGenPln 
      Height          =   495
      Left            =   9285
      TabIndex        =   9
      ToolTipText     =   "Option For Auto-Generating The Plan After The Specified Interval"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Auto Gen &Plan"
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
   Begin MwfCtl.MWCTL_BTN1 CmdRedoPln 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   8340
      TabIndex        =   8
      ToolTipText     =   "Option For Redoing The Plan"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "ReDo P&lan"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDsgDet 
      Height          =   495
      Left            =   7395
      TabIndex        =   7
      ToolTipText     =   "Option For Viewing The Order Design Details"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Ord &Dsg Details"
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
   Begin MwfCtl.MWCTL_BTN1 CmdShwLd 
      Height          =   495
      Left            =   6450
      TabIndex        =   6
      ToolTipText     =   "Option For Displaying The Load/ Capacity Utilization"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Show  &Cap/Load"
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
   Begin MwfCtl.MWCTL_BTN1 CmdGenPln 
      Height          =   495
      Left            =   5505
      TabIndex        =   5
      ToolTipText     =   "Option For Generating The Plan"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "G&en Plan"
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
   Begin MwfCtl.MWCTL_BTN1 CmdUpdStat 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Option For Updating The Plan Order Status"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Update Status"
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
   Begin MwfCtl.MWCTL_BTN1 CmdResPln 
      Height          =   495
      Left            =   3615
      TabIndex        =   3
      ToolTipText     =   "Option For Deleting The Order Plan Date Records"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Reset Plan"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSC 
      Height          =   495
      Left            =   2670
      TabIndex        =   2
      ToolTipText     =   "Option To Select The Scope Options For The Entry"
      Top             =   9270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   873
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Scope"
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
      Left            =   0
      TabIndex        =   135
      Top             =   9300
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9630
      Left            =   -120
      TabIndex        =   136
      Top             =   -60
      Width           =   14145
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   150
         TabIndex        =   138
         Top             =   120
         Width           =   13500
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            TabIndex        =   140
            Top             =   5700
            Width           =   5415
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   1080
               TabIndex        =   141
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WOPWIPP1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   1920
               TabIndex        =   142
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WOPWIPP2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   2760
               TabIndex        =   143
               Top             =   0
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WOPWIPP3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   3570
               TabIndex        =   144
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WOPWIPP4"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   4410
               TabIndex        =   145
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "####0.0"
               MaxLength       =   7
               DataType        =   2
               IdName          =   "WOPWIPP5"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "WIP"
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
               Left            =   180
               TabIndex        =   153
               Top             =   0
               Width           =   675
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5460
            TabIndex        =   146
            Top             =   5700
            Width           =   6285
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   870
               TabIndex        =   147
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdOp(WP1LDPPC)"
               IdName          =   "WOPDLDP1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   1890
               TabIndex        =   148
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdOp(WP2LDPPC)"
               IdName          =   "WOPDLDP2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   2910
               TabIndex        =   149
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdOp(WP3LDPPC)"
               IdName          =   "WOPDLDP3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   3930
               TabIndex        =   150
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdOp(WP4LDPPC)"
               IdName          =   "WOPDLDP4"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   4950
               TabIndex        =   151
               Top             =   0
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               Enabled         =   0   'False
               Mask            =   "######0.0"
               MaxLength       =   9
               DataType        =   2
               ReCalcOn        =   "GrdOp(WP5LDPPC)"
               IdName          =   "WOPDLDP5"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Load/Pc"
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
               Left            =   30
               TabIndex        =   152
               Top             =   0
               Width           =   915
            End
         End
         Begin MwfCtl.MWCTL_FLX GrdOp 
            Height          =   5700
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   12945
            _ExtentX        =   22834
            _ExtentY        =   10054
            Cols            =   26
            colname1        =   "OPTC"
            heading1        =   "Tc   "
            datafld1        =   "OpTc"
            datatype1       =   4
            recalcparent1   =   $"EmrFrmPpcOrdPln.frx":0000
            maxlength1      =   3
            colname2        =   "OPYY"
            heading2        =   "Yy"
            datafld2        =   "OpYy"
            datatype2       =   4
            recalcparent2   =   $"EmrFrmPpcOrdPln.frx":00A5
            maxlength2      =   2
            colname3        =   "OPCHR"
            heading3        =   "Chr    "
            datafld3        =   "OpChr"
            datatype3       =   4
            recalcparent3   =   $"EmrFrmPpcOrdPln.frx":014A
            maxlength3      =   3
            colname4        =   "OPNO"
            heading4        =   "No          "
            datafld4        =   "OpNo"
            datatype4       =   1
            mask4           =   "######0"
            recalcparent4   =   $"EmrFrmPpcOrdPln.frx":01EF
            maxlength4      =   7
            colname5        =   "OPSR"
            heading5        =   "Sr     "
            datafld5        =   "OpSr"
            datatype5       =   1
            mask5           =   "###0"
            recalcparent5   =   $"EmrFrmPpcOrdPln.frx":0294
            maxlength5      =   4
            colname6        =   "WODDELDT"
            heading6        =   "Del Dt     "
            datatype6       =   3
            recalcon6       =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength6      =   8
            style6          =   2
            colname7        =   "OPLINE"
            heading7        =   "Line     "
            datafld7        =   "OpLine"
            datatype7       =   4
            recalcon7       =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength7      =   5
            colname8        =   "OPWDIAAVLBLDT"
            heading8        =   "wDiaDt    "
            datafld8        =   "OpWDiaAvlblDt"
            datatype8       =   3
            recalcon8       =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength8      =   8
            colname9        =   "OPHDIAAVLBLDT"
            heading9        =   "hDiaDt    "
            datafld9        =   "OpHDiaAvlblDt"
            datatype9       =   3
            recalcon9       =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength9      =   8
            colname10       =   "WODDMCD"
            heading10       =   "Design Code              "
            datatype10      =   4
            recalcparent10  =   "GrdOp(WP1LDPPC)"
            maxlength10     =   15
            style10         =   2
            colname11       =   "WODWH"
            heading11       =   "W/ H"
            datatype11      =   4
            recalcparent11  =   "GrdOp(WP1LDPPC)"
            maxlength11     =   2
            style11         =   2
            colname12       =   "WODPRTCD"
            heading12       =   "Prt"
            datatype12      =   4
            maxlength12     =   2
            style12         =   2
            colname13       =   "WODPRDQTY"
            heading13       =   "Prd Qty"
            datatype13      =   2
            mask13          =   "####0.0"
            maxlength13     =   7
            style13         =   2
            colname14       =   "WBALQTY"
            heading14       =   "Bal Qty "
            datatype14      =   2
            mask14          =   "####0.0"
            maxlength14     =   7
            style14         =   2
            colname15       =   "WFLRQTY"
            heading15       =   "Flr Qty  "
            datatype15      =   2
            mask15          =   "####0.0"
            recalcon15      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength15     =   7
            style15         =   2
            colname16       =   "OPWIPP1"
            datafld16       =   "OpWipP1"
            datatype16      =   2
            mask16          =   "####0.0"
            recalcon16      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength16     =   7
            style16         =   0
            colname17       =   "OPWIPP2"
            datafld17       =   "OpWipP2"
            datatype17      =   2
            mask17          =   "####0.0"
            recalcon17      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength17     =   7
            style17         =   0
            colname18       =   "OPWIPP3"
            datafld18       =   "OpWipP3"
            datatype18      =   2
            mask18          =   "####0.0"
            recalcon18      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength18     =   7
            style18         =   0
            colname19       =   "OPWIPP4"
            datafld19       =   "OpWipP4"
            datatype19      =   2
            mask19          =   "####0.0"
            recalcon19      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength19     =   7
            style19         =   0
            colname20       =   "OPWIPP5"
            datafld20       =   "OpWipP5"
            datatype20      =   2
            mask20          =   "####0.0"
            recalcon20      =   "GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)"
            maxlength20     =   7
            style20         =   0
            colname21       =   "WP1LDPPC"
            datatype21      =   2
            mask21          =   "######0.0"
            recalcon21      =   "GrdOp(WODDMCD),GrdOp(WODWH)"
            recalcparent21  =   "WOPDLDP1"
            maxlength21     =   9
            style21         =   0
            colname22       =   "WP2LDPPC"
            datatype22      =   2
            mask22          =   "######0.0"
            recalcparent22  =   "WOPDLDP2"
            maxlength22     =   9
            style22         =   0
            colname23       =   "WP3LDPPC"
            datatype23      =   2
            mask23          =   "######0.0"
            recalcparent23  =   "WOPDLDP3"
            maxlength23     =   9
            style23         =   0
            colname24       =   "WP4LDPPC"
            datatype24      =   2
            mask24          =   "######0.0"
            recalcparent24  =   "WOPDLDP4"
            maxlength24     =   9
            style24         =   0
            colname25       =   "WP5LDPPC"
            datatype25      =   2
            mask25          =   "######0.0"
            recalcparent25  =   "WOPDLDP5"
            maxlength25     =   9
            style25         =   0
         End
         Begin MwfCtl.MWCTL_FLX GrdOpd 
            Height          =   3120
            Left            =   0
            TabIndex        =   1
            Top             =   6030
            Width           =   12945
            _ExtentX        =   22834
            _ExtentY        =   5503
            Cols            =   13
            colname1        =   "OPDPLNDT"
            heading1        =   "Pln  Dt     "
            datafld1        =   "OpdPlnDt"
            datatype1       =   3
            maxlength1      =   8
            colname2        =   "OPDP1"
            heading2        =   "P1          "
            datafld2        =   "OpdP1"
            datatype2       =   2
            mask2           =   "####0.0"
            recalcparent2   =   "GrdOpd(WOPDLOADP1)"
            maxlength2      =   7
            colname3        =   "OPDP2"
            heading3        =   "P2          "
            datafld3        =   "OpdP2"
            datatype3       =   2
            mask3           =   "####0.0"
            recalcparent3   =   "GrdOpd(WOPDLOADP2)"
            maxlength3      =   7
            colname4        =   "OPDP3"
            heading4        =   "P3          "
            datafld4        =   "OpdP3"
            datatype4       =   2
            mask4           =   "####0.0"
            recalcparent4   =   "GrdOpd(WOPDLOADP3)"
            maxlength4      =   7
            colname5        =   "OPDP4"
            heading5        =   "P4          "
            datafld5        =   "OpdP4"
            datatype5       =   2
            mask5           =   "####0.0"
            recalcparent5   =   "GrdOpd(WOPDLOADP4)"
            maxlength5      =   7
            colname6        =   "OPDP5"
            heading6        =   "P5          "
            datafld6        =   "OpdP5"
            datatype6       =   2
            mask6           =   "####0.0"
            recalcparent6   =   "GrdOpd(WOPDLOADP5)"
            maxlength6      =   7
            colname7        =   "WLOAD"
            heading7        =   "                       "
            datatype7       =   4
            maxlength7      =   1
            style7          =   2
            colname8        =   "WOPDLOADP1"
            heading8        =   "P1               "
            datatype8       =   2
            mask8           =   "######0.0"
            recalcon8       =   "GrdOpd(OPDP1)"
            maxlength8      =   9
            style8          =   2
            colname9        =   "WOPDLOADP2"
            heading9        =   "P2              "
            datatype9       =   2
            mask9           =   "######0.0"
            recalcon9       =   "GrdOpd(OPDP2)"
            maxlength9      =   9
            style9          =   2
            colname10       =   "WOPDLOADP3"
            heading10       =   "P3              "
            datatype10      =   2
            mask10          =   "######0.0"
            recalcon10      =   "GrdOpd(OPDP3)"
            maxlength10     =   9
            style10         =   2
            colname11       =   "WOPDLOADP4"
            heading11       =   "P4              "
            datatype11      =   2
            mask11          =   "######0.0"
            recalcon11      =   "GrdOpd(OPDP4)"
            maxlength11     =   9
            style11         =   2
            colname12       =   "WOPDLOADP5"
            heading12       =   "P5              "
            datatype12      =   2
            mask12          =   "######0.0"
            recalcon12      =   "GrdOpd(OPDP5)"
            maxlength12     =   9
            style12         =   2
         End
      End
      Begin VB.Frame FraUpdStat 
         Height          =   3375
         Left            =   3750
         TabIndex        =   184
         Top             =   5895
         Visible         =   0   'False
         Width           =   5595
         Begin MwfCtl.MWCTL_BTN1 CmdUpdStatGo 
            Height          =   285
            Left            =   4800
            TabIndex        =   73
            ToolTipText     =   "Click Here To Update The Order Plan Status"
            Top             =   2805
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
            Index           =   47
            Left            =   1770
            TabIndex        =   52
            ToolTipText     =   "Enter From Order Tc"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPTCFR"
            CmpStr          =   "OdTc >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   48
            Left            =   2445
            TabIndex        =   53
            ToolTipText     =   "Enter From Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPYYFR"
            CmpStr          =   "OdYy >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   49
            Left            =   2880
            TabIndex        =   54
            ToolTipText     =   "Enter From Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPCHRFR"
            CmpStr          =   "OdChr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   50
            Left            =   1770
            TabIndex        =   58
            ToolTipText     =   "Enter From Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WUSOPNOFR"
            CmpStr          =   "OdNo >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   51
            Left            =   2700
            TabIndex        =   59
            ToolTipText     =   "Enter From Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WUSOPSRFR"
            CmpStr          =   "OdSr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   52
            Left            =   3750
            TabIndex        =   55
            ToolTipText     =   "Enter To Order Tc"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPTCTO"
            CmpStr          =   "OdTc <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   53
            Left            =   4395
            TabIndex        =   56
            ToolTipText     =   "Enter To Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPYYTO"
            CmpStr          =   "OdYy <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   54
            Left            =   4830
            TabIndex        =   57
            ToolTipText     =   "Enter To Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOPCHRTO"
            CmpStr          =   "OdChr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   55
            Left            =   3750
            TabIndex        =   60
            ToolTipText     =   "Enter To Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WUSOPNOTO"
            CmpStr          =   "OdNo <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   56
            Left            =   4680
            TabIndex        =   61
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WUSOPSRTO"
            CmpStr          =   "OdSr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   57
            Left            =   1770
            TabIndex        =   64
            ToolTipText     =   "Enter From Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOMCMCDFR"
            CmpStr          =   "OmCmCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   58
            Left            =   3750
            TabIndex        =   65
            ToolTipText     =   "Enter To Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WUSOMCMCDTO"
            CmpStr          =   "OmCmCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   59
            Left            =   1770
            TabIndex        =   66
            ToolTipText     =   "Enter From Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WUSOMDTFR"
            CmpStr          =   "OmDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   60
            Left            =   3750
            TabIndex        =   67
            ToolTipText     =   "Enter To Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WUSOMDTTO"
            CmpStr          =   "OmDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   61
            Left            =   1770
            TabIndex        =   70
            ToolTipText     =   "Enter From Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WUSODPRTCDFR"
            CmpStr          =   "OdPrtCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   62
            Left            =   3750
            TabIndex        =   71
            ToolTipText     =   "Enter To Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   2
            DataType        =   4
            IdName          =   "WUSODPRTCDTO"
            CmpStr          =   "OdPrtCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   63
            Left            =   1770
            TabIndex        =   68
            ToolTipText     =   "Enter From Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WUSODDELDTFR"
            CmpStr          =   "OdDelDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   64
            Left            =   3750
            TabIndex        =   69
            ToolTipText     =   "Enter To Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WUSODDELDTTO"
            CmpStr          =   "OdDelDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   118
            Left            =   1770
            TabIndex        =   62
            ToolTipText     =   "Enter From Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WUSOPLINEFR"
            CmpStr          =   "OdLine >="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   119
            Left            =   3750
            TabIndex        =   63
            ToolTipText     =   "Enter To Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WUSOPLINETO"
            CmpStr          =   "OdLine <="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   120
            Left            =   1770
            TabIndex        =   72
            ToolTipText     =   $"EmrFrmPpcOrdPln.frx":0339
            Top             =   2805
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WUSGETNEWORD"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Get New Orders"
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
            Index           =   79
            Left            =   120
            TabIndex        =   268
            Top             =   2805
            Width           =   1605
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
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
            Index           =   78
            Left            =   120
            TabIndex        =   267
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Update Status"
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
            TabIndex        =   262
            Top             =   90
            Width           =   5595
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
            Index           =   43
            Left            =   3750
            TabIndex        =   198
            Top             =   390
            Width           =   675
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
            Index           =   42
            Left            =   1770
            TabIndex        =   197
            Top             =   390
            Width           =   675
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
            Index           =   41
            Left            =   4590
            TabIndex        =   196
            Top             =   1095
            Width           =   165
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
            Index           =   40
            Left            =   4740
            TabIndex        =   195
            Top             =   810
            Width           =   165
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
            Index           =   39
            Left            =   4320
            TabIndex        =   194
            Top             =   810
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord No/ Sr"
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
            Left            =   120
            TabIndex        =   193
            Top             =   1095
            Width           =   1185
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
            Index           =   37
            Left            =   2610
            TabIndex        =   192
            Top             =   1095
            Width           =   165
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
            Index           =   36
            Left            =   2790
            TabIndex        =   191
            Top             =   810
            Width           =   165
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   2280
            TabIndex        =   190
            Top             =   810
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Tc/ Yy/ Chr"
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
            Left            =   120
            TabIndex        =   189
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Left            =   120
            TabIndex        =   188
            Top             =   1665
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Dt"
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
            Left            =   120
            TabIndex        =   187
            Top             =   1950
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority"
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
            Left            =   120
            TabIndex        =   186
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Del Dt"
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
            Left            =   120
            TabIndex        =   185
            Top             =   2235
            Width           =   1095
         End
      End
      Begin VB.Frame FraShwLd 
         Height          =   3855
         Left            =   150
         TabIndex        =   234
         Top             =   5415
         Visible         =   0   'False
         Width           =   12945
         Begin MwfCtl.MWCTL_BTN1 CmdShwLdGo 
            Height          =   285
            Left            =   1560
            TabIndex        =   78
            ToolTipText     =   "Click Here To Generate Plan"
            Top             =   1365
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
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
            Index           =   103
            Left            =   1230
            TabIndex        =   76
            ToolTipText     =   "Enter From Plan Date"
            Top             =   735
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WSLPLNDTFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   104
            Left            =   1230
            TabIndex        =   77
            ToolTipText     =   "Enter To Plan Date"
            Top             =   1020
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WSLPLNDTTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   92
            Left            =   1230
            TabIndex        =   75
            ToolTipText     =   "Enter Line"
            Top             =   450
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WSLLINE"
         End
         Begin MwfCtl.MWCTL_FLX GrdSL 
            Height          =   2400
            Left            =   2310
            TabIndex        =   79
            Top             =   420
            Width           =   7650
            _ExtentX        =   13494
            _ExtentY        =   4233
            Cols            =   17
            colname1        =   "WSLDT"
            heading1        =   "Date                 "
            datatype1       =   3
            maxlength1      =   8
            tooltiptext1    =   "Plan Date"
            colname2        =   "WSLP1"
            heading2        =   "P1                 "
            datatype2       =   2
            mask2           =   "#######0.0"
            maxlength2      =   10
            tooltiptext2    =   "P1 pcs"
            colname3        =   "WSLP2"
            heading3        =   "P2                 "
            datatype3       =   2
            mask3           =   "#######0.0"
            maxlength3      =   10
            tooltiptext3    =   "P2 Pcs"
            colname4        =   "WSLP3"
            heading4        =   "P3                 "
            datatype4       =   2
            mask4           =   "#######0.0"
            maxlength4      =   10
            tooltiptext4    =   "P3 Pcs"
            colname5        =   "WSLP4"
            heading5        =   "P4                "
            datatype5       =   2
            mask5           =   "#######0.0"
            maxlength5      =   10
            tooltiptext5    =   "P4 Pcs"
            colname6        =   "WSLP5"
            heading6        =   "P5                 "
            datatype6       =   2
            mask6           =   "#######0.0"
            maxlength6      =   10
            tooltiptext6    =   "P5 Pcs"
            colname7        =   "WSLP1LD"
            datatype7       =   1
            mask7           =   "#########0"
            maxlength7      =   10
            style7          =   0
            colname8        =   "WSLP2LD"
            datatype8       =   1
            mask8           =   "#########0"
            maxlength8      =   10
            style8          =   0
            colname9        =   "WSLP3LD"
            datatype9       =   1
            mask9           =   "#########0"
            maxlength9      =   10
            style9          =   0
            colname10       =   "WSLP4LD"
            datatype10      =   1
            mask10          =   "#########0"
            maxlength10     =   10
            style10         =   0
            colname11       =   "WSLP5LD"
            datatype11      =   1
            mask11          =   "#########0"
            maxlength11     =   10
            style11         =   0
            colname12       =   "WSLP1LDALL"
            datatype12      =   1
            mask12          =   "#########0"
            maxlength12     =   10
            style12         =   0
            colname13       =   "WSLP2LDALL"
            datatype13      =   1
            mask13          =   "#########0"
            maxlength13     =   10
            style13         =   0
            colname14       =   "WSLP3LDALL"
            datatype14      =   1
            mask14          =   "#########0"
            maxlength14     =   10
            style14         =   0
            colname15       =   "WSLP4LDALL"
            datatype15      =   1
            mask15          =   "#########0"
            maxlength15     =   10
            style15         =   0
            colname16       =   "WSLP5LDALL"
            datatype16      =   1
            mask16          =   "#########0"
            maxlength16     =   10
            style16         =   0
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   94
            Left            =   3720
            TabIndex        =   238
            Top             =   2820
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.0"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WSLLINELDP1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   95
            Left            =   7230
            TabIndex        =   240
            Top             =   2820
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.0"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WSLLINELDP4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   96
            Left            =   8400
            TabIndex        =   241
            Top             =   2820
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.0"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WSLLINELDP5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   97
            Left            =   6060
            TabIndex        =   242
            Top             =   2820
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.0"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WSLLINELDP3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   98
            Left            =   4890
            TabIndex        =   243
            Top             =   2820
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "#######0.0"
            MaxLength       =   10
            DataType        =   2
            IdName          =   "WSLLINELDP2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   99
            Left            =   3720
            TabIndex        =   244
            Top             =   3105
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPP1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   100
            Left            =   7230
            TabIndex        =   245
            Top             =   3105
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPP4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   101
            Left            =   8400
            TabIndex        =   246
            Top             =   3105
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPP5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   102
            Left            =   6060
            TabIndex        =   247
            Top             =   3105
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPP3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   105
            Left            =   4890
            TabIndex        =   248
            Top             =   3105
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPP2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   106
            Left            =   3720
            TabIndex        =   250
            Top             =   3390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPALLP1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   107
            Left            =   7230
            TabIndex        =   251
            Top             =   3390
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPALLP4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   108
            Left            =   8400
            TabIndex        =   252
            Top             =   3390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPALLP5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   109
            Left            =   6060
            TabIndex        =   253
            Top             =   3390
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPALLP3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   110
            Left            =   4890
            TabIndex        =   254
            Top             =   3390
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            Alignment       =   2
            Enabled         =   0   'False
            Mask            =   "-#######0.0"
            MaxLength       =   11
            DataType        =   2
            IdName          =   "WSLBALCAPALLP2"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Show Load"
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
            TabIndex        =   257
            Top             =   90
            Width           =   12975
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Cap All"
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
            Left            =   2310
            TabIndex        =   255
            Top             =   3390
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Line Cap"
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
            Left            =   2310
            TabIndex        =   249
            Top             =   3105
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Load"
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
            Left            =   2310
            TabIndex        =   239
            Top             =   2820
            Width           =   1395
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
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
            Index           =   67
            Left            =   120
            TabIndex        =   237
            Top             =   450
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan Dt To"
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
            Index           =   87
            Left            =   120
            TabIndex        =   236
            Top             =   1020
            Width           =   1185
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan Dt Fr"
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
            Left            =   120
            TabIndex        =   235
            Top             =   735
            Width           =   1185
         End
      End
      Begin VB.Frame FraResPln 
         Height          =   3105
         Left            =   3780
         TabIndex        =   169
         Top             =   6165
         Visible         =   0   'False
         Width           =   5625
         Begin MwfCtl.MWCTL_BTN1 CmdResPlnGo 
            Height          =   285
            Left            =   4800
            TabIndex        =   51
            ToolTipText     =   "Click Here To Reset Plan"
            Top             =   2550
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
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
            Index           =   0
            Left            =   1770
            TabIndex        =   31
            ToolTipText     =   "Enter From Order Tc"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPTCFR"
            CmpStr          =   "OdTc >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   7
            Left            =   2445
            TabIndex        =   32
            ToolTipText     =   "Enter From Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPYYFR"
            CmpStr          =   "OdYy >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   13
            Left            =   2880
            TabIndex        =   33
            ToolTipText     =   "Enter From Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPCHRFR"
            CmpStr          =   "OdChr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   1770
            TabIndex        =   37
            ToolTipText     =   "Enter From Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WRPOPNOFR"
            CmpStr          =   "OdNo >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   15
            Left            =   2700
            TabIndex        =   38
            ToolTipText     =   "Enter From Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WRPOPSRFR"
            CmpStr          =   "OdSr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   3750
            TabIndex        =   34
            ToolTipText     =   "Enter To Order Tc"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPTCTO"
            CmpStr          =   "OdTc <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   17
            Left            =   4395
            TabIndex        =   35
            ToolTipText     =   "Enter To Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPYYTO"
            CmpStr          =   "OdYy <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   18
            Left            =   4830
            TabIndex        =   36
            ToolTipText     =   "Enter To Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOPCHRTO"
            CmpStr          =   "OdChr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   19
            Left            =   3750
            TabIndex        =   39
            ToolTipText     =   "Enter To Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WRPOPNOTO"
            CmpStr          =   "OdNo <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   20
            Left            =   4680
            TabIndex        =   40
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WRPOPSRTO"
            CmpStr          =   "OdSr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   21
            Left            =   1770
            TabIndex        =   43
            ToolTipText     =   "Enter From Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOMCMCDFR"
            CmpStr          =   "OmCmCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   22
            Left            =   3750
            TabIndex        =   44
            ToolTipText     =   "Enter To Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPOMCMCDTO"
            CmpStr          =   "OmCmCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   41
            Left            =   1770
            TabIndex        =   45
            ToolTipText     =   "Enter From Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WRPOMDTFR"
            CmpStr          =   "OmDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   42
            Left            =   3750
            TabIndex        =   46
            ToolTipText     =   "Enter To Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WRPOMDTTO"
            CmpStr          =   "OmDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   43
            Left            =   1770
            TabIndex        =   49
            ToolTipText     =   "Enter From Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPODPRTCDFR"
            CmpStr          =   "OdPrtCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   44
            Left            =   3750
            TabIndex        =   50
            ToolTipText     =   "Enter To Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WRPODPRTCDTO"
            CmpStr          =   "OdPrtCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   45
            Left            =   1770
            TabIndex        =   47
            ToolTipText     =   "Enter From Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WRPODDELDTFR"
            CmpStr          =   "OdDelDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   46
            Left            =   3750
            TabIndex        =   48
            ToolTipText     =   "Enter To Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WRPODDELDTTO"
            CmpStr          =   "OdDelDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   116
            Left            =   1770
            TabIndex        =   41
            ToolTipText     =   "Enter From Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WRPOPLINEFR"
            CmpStr          =   "OdLine >="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   117
            Left            =   3750
            TabIndex        =   42
            ToolTipText     =   "Enter To Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WRPOPLINETO"
            CmpStr          =   "OdLine <="
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
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
            Left            =   120
            TabIndex        =   266
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Reset Plan"
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
            TabIndex        =   261
            Top             =   90
            Width           =   5655
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Del Dt"
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
            Left            =   120
            TabIndex        =   183
            Top             =   2235
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority"
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
            Left            =   120
            TabIndex        =   182
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Dt"
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
            Left            =   120
            TabIndex        =   181
            Top             =   1950
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Left            =   120
            TabIndex        =   180
            Top             =   1665
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Tc/ Yy/ Chr"
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
            Left            =   120
            TabIndex        =   179
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   2280
            TabIndex        =   178
            Top             =   810
            Width           =   195
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
            Index           =   16
            Left            =   2790
            TabIndex        =   177
            Top             =   810
            Width           =   165
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
            Index           =   11
            Left            =   2610
            TabIndex        =   176
            Top             =   1095
            Width           =   165
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord No/ Sr"
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
            Left            =   120
            TabIndex        =   175
            Top             =   1095
            Width           =   1185
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
            Index           =   8
            Left            =   4320
            TabIndex        =   174
            Top             =   810
            Width           =   105
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
            Index           =   7
            Left            =   4740
            TabIndex        =   173
            Top             =   810
            Width           =   165
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
            Index           =   6
            Left            =   4590
            TabIndex        =   172
            Top             =   1095
            Width           =   165
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
            Index           =   5
            Left            =   1770
            TabIndex        =   171
            Top             =   390
            Width           =   675
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
            Index           =   4
            Left            =   3750
            TabIndex        =   170
            Top             =   390
            Width           =   675
         End
      End
      Begin VB.Frame FraRedoPln 
         Height          =   1155
         Left            =   3720
         TabIndex        =   232
         Top             =   8115
         Visible         =   0   'False
         Width           =   5655
         Begin MwfCtl.MWCTL_BTN1 CmdRedoPlnGo 
            Height          =   285
            Left            =   4800
            TabIndex        =   74
            ToolTipText     =   "Click Here To Generate Plan"
            Top             =   570
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Redo Plan"
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
            TabIndex        =   259
            Top             =   90
            Width           =   5655
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Are You Sure You Want To Redo The Plan?"
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
            Index           =   77
            Left            =   150
            TabIndex        =   233
            Top             =   570
            Width           =   4575
         End
      End
      Begin VB.Frame FraGenPln 
         Height          =   3405
         Left            =   3720
         TabIndex        =   199
         Top             =   5865
         Visible         =   0   'False
         Width           =   5685
         Begin MwfCtl.MWCTL_BTN1 CmdGenPlnGo 
            Height          =   285
            Left            =   4950
            TabIndex        =   133
            ToolTipText     =   "Click Here To Generate Plan"
            Top             =   2805
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
            Index           =   65
            Left            =   1770
            TabIndex        =   90
            ToolTipText     =   "Enter From Order Tc "
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPTCFR"
            CmpStr          =   "OdTc>= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   66
            Left            =   2445
            TabIndex        =   91
            ToolTipText     =   "Enter From Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPYYFR"
            CmpStr          =   "OdYy>= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   67
            Left            =   2880
            TabIndex        =   92
            ToolTipText     =   "Enter From Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPCHRFR"
            CmpStr          =   "OdChr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   68
            Left            =   1770
            TabIndex        =   96
            ToolTipText     =   "Enter From Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WGPOPNOFR"
            CmpStr          =   "OdNo >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   69
            Left            =   2700
            TabIndex        =   97
            ToolTipText     =   "Enter From Order Serial "
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WGPOPSRFR"
            CmpStr          =   "OdSr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   70
            Left            =   3750
            TabIndex        =   93
            ToolTipText     =   "Enter To Order Tc "
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPTCTO"
            CmpStr          =   "OdTc <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   71
            Left            =   4395
            TabIndex        =   94
            ToolTipText     =   "Enter To Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPYYTO"
            CmpStr          =   "OdYy <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   72
            Left            =   4830
            TabIndex        =   95
            ToolTipText     =   "Enter To Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOPCHRTO"
            CmpStr          =   "OdChr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   73
            Left            =   3750
            TabIndex        =   98
            ToolTipText     =   "Enter To Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WGPOPNOTO"
            CmpStr          =   "OdNo <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   74
            Left            =   4680
            TabIndex        =   117
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WGPOPSRTO"
            CmpStr          =   "OdSr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   75
            Left            =   1770
            TabIndex        =   120
            ToolTipText     =   "Enter From Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOMCMCDFR"
            CmpStr          =   "OmCmCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   76
            Left            =   3750
            TabIndex        =   121
            ToolTipText     =   "Enter To Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPOMCMCDTO"
            CmpStr          =   "OmCmCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   77
            Left            =   1770
            TabIndex        =   122
            ToolTipText     =   "Enter From Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WGPOMDTFR"
            CmpStr          =   "OmDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   78
            Left            =   3750
            TabIndex        =   123
            ToolTipText     =   "Enter To Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WGPOMDTTO"
            CmpStr          =   "OmDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   79
            Left            =   1770
            TabIndex        =   126
            ToolTipText     =   "Enter From Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPODPRTCDFR"
            CmpStr          =   "OdPrtCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   80
            Left            =   3750
            TabIndex        =   127
            ToolTipText     =   "Enter To Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WGPODPRTCDTO"
            CmpStr          =   "OdPrtCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   81
            Left            =   1770
            TabIndex        =   124
            ToolTipText     =   "Enter From Production Delivery Date "
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WGPODDELDTFR"
            CmpStr          =   "OdDelDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   82
            Left            =   3750
            TabIndex        =   125
            ToolTipText     =   "Enter To Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WGPODDELDTTO"
            CmpStr          =   "OdDelDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   83
            Left            =   1770
            TabIndex        =   128
            ToolTipText     =   "Enter Gap Days For P1"
            Top             =   2805
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WGPGAPDYSP1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   84
            Left            =   2415
            TabIndex        =   129
            ToolTipText     =   "Enter Gap Days For P2"
            Top             =   2805
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WGPGAPDYSP2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   85
            Left            =   3060
            TabIndex        =   130
            ToolTipText     =   "Enter Gap Days For P3"
            Top             =   2805
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WGPGAPDYSP3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   86
            Left            =   3690
            TabIndex        =   131
            ToolTipText     =   "Enter Gap Days For P4"
            Top             =   2805
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WGPGAPDYSP4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   87
            Left            =   4350
            TabIndex        =   132
            ToolTipText     =   "Enter Gap Days For P5"
            Top             =   2805
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WGPGAPDYSP5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   114
            Left            =   1770
            TabIndex        =   118
            ToolTipText     =   "Enter From Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WGPOPLINEFR"
            CmpStr          =   "OdLine >="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   115
            Left            =   3750
            TabIndex        =   119
            ToolTipText     =   "Enter To Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WGPOPLINETO"
            CmpStr          =   "OdLine <="
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
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
            Left            =   120
            TabIndex        =   265
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Generate Plan"
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
            TabIndex        =   260
            Top             =   90
            Width           =   5670
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   4185
            TabIndex        =   218
            Top             =   2805
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   3540
            TabIndex        =   217
            Top             =   2805
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   2895
            TabIndex        =   216
            Top             =   2805
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   2250
            TabIndex        =   215
            Top             =   2805
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gap Days"
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
            Left            =   120
            TabIndex        =   214
            Top             =   2805
            Width           =   1455
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Del Dt"
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
            Left            =   120
            TabIndex        =   213
            Top             =   2235
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority"
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
            Left            =   120
            TabIndex        =   212
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Dt"
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
            Left            =   120
            TabIndex        =   211
            Top             =   1950
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Left            =   120
            TabIndex        =   210
            Top             =   1665
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Tc/ Yy/ Chr"
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
            Left            =   120
            TabIndex        =   209
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Left            =   2280
            TabIndex        =   208
            Top             =   810
            Width           =   195
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
            Index           =   51
            Left            =   2790
            TabIndex        =   207
            Top             =   810
            Width           =   165
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
            Index           =   50
            Left            =   2610
            TabIndex        =   206
            Top             =   1095
            Width           =   165
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord No/ Sr"
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
            Left            =   120
            TabIndex        =   205
            Top             =   1095
            Width           =   1185
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
            Index           =   48
            Left            =   4320
            TabIndex        =   204
            Top             =   810
            Width           =   105
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
            Index           =   47
            Left            =   4740
            TabIndex        =   203
            Top             =   810
            Width           =   165
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
            Index           =   46
            Left            =   4590
            TabIndex        =   202
            Top             =   1095
            Width           =   165
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
            Index           =   45
            Left            =   1770
            TabIndex        =   201
            Top             =   390
            Width           =   675
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
            Index           =   44
            Left            =   3750
            TabIndex        =   200
            Top             =   390
            Width           =   675
         End
      End
      Begin VB.Frame FraDsgDet 
         Height          =   3135
         Left            =   150
         TabIndex        =   219
         Top             =   6135
         Visible         =   0   'False
         Width           =   12945
         Begin VB.Frame FraPic3D 
            BorderStyle     =   0  'None
            Height          =   1835
            Left            =   5610
            TabIndex        =   227
            Top             =   450
            Width           =   1775
            Begin VB.Image Pic3D 
               Height          =   1770
               Left            =   0
               ToolTipText     =   "3D - Image"
               Top             =   0
               Width           =   1710
            End
         End
         Begin VB.Frame FraPicHsk 
            BorderStyle     =   0  'None
            Height          =   1835
            Left            =   3480
            TabIndex        =   226
            Top             =   450
            Width           =   1775
            Begin VB.Image PicHsk 
               Height          =   1770
               Left            =   0
               ToolTipText     =   "Hand Sketch Image"
               Top             =   0
               Width           =   1710
            End
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   88
            Left            =   1515
            TabIndex        =   220
            ToolTipText     =   "Design Size"
            Top             =   480
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDDMSZ"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   89
            Left            =   1515
            TabIndex        =   221
            ToolTipText     =   "Design Suffix"
            Top             =   765
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   10
            DataType        =   4
            IdName          =   "WDDSFX"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   93
            Left            =   1515
            TabIndex        =   222
            ToolTipText     =   "Worker"
            Top             =   1050
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDKT"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   90
            Left            =   1515
            TabIndex        =   228
            ToolTipText     =   "Design Category"
            Top             =   1335
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDDMCTG"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   91
            Left            =   1515
            TabIndex        =   230
            ToolTipText     =   "Enter Value Addition Category"
            Top             =   1620
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            Enabled         =   0   'False
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WDDVACTG"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Design Details"
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
            TabIndex        =   258
            Top             =   90
            Width           =   12975
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
            Index           =   66
            Left            =   120
            TabIndex        =   231
            Top             =   1620
            Width           =   1395
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
            Index           =   68
            Left            =   120
            TabIndex        =   229
            Top             =   1335
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
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
            Left            =   120
            TabIndex        =   225
            Top             =   480
            Width           =   915
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix"
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
            Left            =   120
            TabIndex        =   224
            Top             =   765
            Width           =   915
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
            Index           =   63
            Left            =   120
            TabIndex        =   223
            Top             =   1050
            Width           =   675
         End
      End
      Begin VB.Frame FraAutoGenPln 
         Height          =   3885
         Left            =   3690
         TabIndex        =   269
         Top             =   5385
         Visible         =   0   'False
         Width           =   5685
         Begin VB.Timer TimerAutoGen 
            Enabled         =   0   'False
            Left            =   5040
            Top             =   2520
         End
         Begin MwfCtl.MWCTL_BTN1 CmdAutoGenPlnGo 
            Height          =   285
            Left            =   4890
            TabIndex        =   116
            ToolTipText     =   "Click Here To Generate Plan"
            Top             =   3375
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
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
            Index           =   121
            Left            =   1890
            TabIndex        =   80
            ToolTipText     =   "Enter From Order Tc "
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPTCFR"
            CmpStr          =   "OdTc>= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   122
            Left            =   2565
            TabIndex        =   81
            ToolTipText     =   "Enter From Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPYYFR"
            CmpStr          =   "OdYy>= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   123
            Left            =   3000
            TabIndex        =   82
            ToolTipText     =   "Enter From Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPCHRFR"
            CmpStr          =   "OdChr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   124
            Left            =   1890
            TabIndex        =   86
            ToolTipText     =   "Enter From Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WAGOPNOFR"
            CmpStr          =   "OdNo >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   125
            Left            =   2820
            TabIndex        =   87
            ToolTipText     =   "Enter From Order Serial "
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WAGOPSRFR"
            CmpStr          =   "OdSr >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   126
            Left            =   3870
            TabIndex        =   83
            ToolTipText     =   "Enter To Order Tc "
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPTCTO"
            CmpStr          =   "OdTc <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   127
            Left            =   4515
            TabIndex        =   84
            ToolTipText     =   "Enter To Order Year"
            Top             =   810
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPYYTO"
            CmpStr          =   "OdYy <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   128
            Left            =   4950
            TabIndex        =   85
            ToolTipText     =   "Enter To Order Character"
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOPCHRTO"
            CmpStr          =   "OdChr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   129
            Left            =   3870
            TabIndex        =   88
            ToolTipText     =   "Enter To Order Number"
            Top             =   1095
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            Mask            =   "######0"
            MaxLength       =   7
            DataType        =   1
            IdName          =   "WAGOPNOTO"
            CmpStr          =   "OdNo <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   130
            Left            =   4800
            TabIndex        =   89
            ToolTipText     =   "Enter To Order Serial"
            Top             =   1095
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "###0"
            MaxLength       =   4
            DataType        =   1
            IdName          =   "WAGOPSRTO"
            CmpStr          =   "OdSr <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   131
            Left            =   1890
            TabIndex        =   101
            ToolTipText     =   "Enter From Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOMCMCDFR"
            CmpStr          =   "OmCmCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   132
            Left            =   3870
            TabIndex        =   102
            ToolTipText     =   "Enter To Customer Code"
            Top             =   1665
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGOMCMCDTO"
            CmpStr          =   "OmCmCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   133
            Left            =   1890
            TabIndex        =   103
            ToolTipText     =   "Enter From Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAGOMDTFR"
            CmpStr          =   "OmDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   134
            Left            =   3870
            TabIndex        =   104
            ToolTipText     =   "Enter To Order Date"
            Top             =   1950
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAGOMDTTO"
            CmpStr          =   "OmDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   135
            Left            =   1890
            TabIndex        =   107
            ToolTipText     =   "Enter From Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGODPRTCDFR"
            CmpStr          =   "OdPrtCd >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   136
            Left            =   3870
            TabIndex        =   108
            ToolTipText     =   "Enter To Priority"
            Top             =   2520
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WAGODPRTCDTO"
            CmpStr          =   "OdPrtCd <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   137
            Left            =   1890
            TabIndex        =   105
            ToolTipText     =   "Enter From Production Delivery Date "
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAGODDELDTFR"
            CmpStr          =   "OdDelDt >= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   138
            Left            =   3870
            TabIndex        =   106
            ToolTipText     =   "Enter To Production Delivery Date"
            Top             =   2235
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   3
            IdName          =   "WAGODDELDTTO"
            CmpStr          =   "OdDelDt <= "
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   139
            Left            =   1890
            TabIndex        =   110
            ToolTipText     =   "Enter Gap Days For P1"
            Top             =   3090
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WAGGAPDYSP1"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   140
            Left            =   2535
            TabIndex        =   111
            ToolTipText     =   "Enter Gap Days For P2"
            Top             =   3090
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WAGGAPDYSP2"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   141
            Left            =   3180
            TabIndex        =   112
            ToolTipText     =   "Enter Gap Days For P3"
            Top             =   3090
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WAGGAPDYSP3"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   142
            Left            =   3810
            TabIndex        =   113
            ToolTipText     =   "Enter Gap Days For P4"
            Top             =   3090
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WAGGAPDYSP4"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   143
            Left            =   4470
            TabIndex        =   114
            ToolTipText     =   "Enter Gap Days For P5"
            Top             =   3090
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Mask            =   "#0"
            MaxLength       =   2
            DataType        =   1
            IdName          =   "WAGGAPDYSP5"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   144
            Left            =   1890
            TabIndex        =   99
            ToolTipText     =   "Enter From Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WAGOPLINEFR"
            CmpStr          =   "OdLine >="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   145
            Left            =   3870
            TabIndex        =   100
            ToolTipText     =   "Enter To Line"
            Top             =   1380
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            MaxLength       =   5
            DataType        =   4
            IdName          =   "WAGOPLINETO"
            CmpStr          =   "OdLine <="
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   146
            Left            =   1890
            TabIndex        =   109
            ToolTipText     =   $"EmrFrmPpcOrdPln.frx":0400
            Top             =   2805
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   503
            DataType        =   4
            IdName          =   "WAGGETNEWORD"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   147
            Left            =   1890
            TabIndex        =   115
            ToolTipText     =   $"EmrFrmPpcOrdPln.frx":04C7
            Top             =   3375
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   503
            Alignment       =   2
            Mask            =   "##0"
            MaxLength       =   3
            DataType        =   1
            IdName          =   "WAGINTR"
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (Minitues)"
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
            Index           =   102
            Left            =   120
            TabIndex        =   292
            Top             =   3375
            Width           =   1725
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Get New Orders"
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
            Index           =   101
            Left            =   120
            TabIndex        =   291
            Top             =   2805
            Width           =   1605
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
            Index           =   100
            Left            =   3750
            TabIndex        =   290
            Top             =   390
            Width           =   675
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
            Index           =   99
            Left            =   1770
            TabIndex        =   289
            Top             =   390
            Width           =   675
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
            Index           =   98
            Left            =   4710
            TabIndex        =   288
            Top             =   1095
            Width           =   165
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
            Index           =   97
            Left            =   4860
            TabIndex        =   287
            Top             =   810
            Width           =   165
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
            Index           =   96
            Left            =   4440
            TabIndex        =   286
            Top             =   810
            Width           =   105
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord No/ Sr"
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
            Index           =   95
            Left            =   120
            TabIndex        =   285
            Top             =   1095
            Width           =   1185
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
            Index           =   94
            Left            =   2730
            TabIndex        =   284
            Top             =   1095
            Width           =   165
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
            Index           =   93
            Left            =   2910
            TabIndex        =   283
            Top             =   810
            Width           =   165
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Index           =   92
            Left            =   2400
            TabIndex        =   282
            Top             =   810
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Tc/ Yy/ Chr"
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
            Index           =   91
            Left            =   120
            TabIndex        =   281
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Index           =   90
            Left            =   120
            TabIndex        =   280
            Top             =   1665
            Width           =   1035
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Ord Dt"
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
            Index           =   89
            Left            =   120
            TabIndex        =   279
            Top             =   1830
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority"
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
            Index           =   88
            Left            =   120
            TabIndex        =   278
            Top             =   2400
            Width           =   795
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Prd Del Dt"
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
            Index           =   86
            Left            =   120
            TabIndex        =   277
            Top             =   2235
            Width           =   1095
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Gap Days"
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
            Index           =   85
            Left            =   120
            TabIndex        =   276
            Top             =   3090
            Width           =   1455
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Index           =   84
            Left            =   2370
            TabIndex        =   275
            Top             =   3090
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Index           =   83
            Left            =   3015
            TabIndex        =   274
            Top             =   3090
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Index           =   82
            Left            =   3660
            TabIndex        =   273
            Top             =   3090
            Width           =   195
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   " /"
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
            Index           =   81
            Left            =   4305
            TabIndex        =   272
            Top             =   3090
            Width           =   195
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Auto Generate Plan"
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
            TabIndex        =   271
            Top             =   90
            Width           =   5670
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Line"
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
            Index           =   80
            Left            =   120
            TabIndex        =   270
            Top             =   1380
            Width           =   1035
         End
      End
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   -36
      TabIndex        =   137
      Top             =   -180
      Width           =   7200
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   435
         Index           =   5
         Left            =   480
         TabIndex        =   139
         ToolTipText     =   "This field is put on the entry form because there should be atleast one atxt field in the form"
         Top             =   30
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   767
         Enabled         =   0   'False
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WATXT"
      End
   End
   Begin VB.Frame FraSC 
      Height          =   3405
      Left            =   3600
      TabIndex        =   154
      Top             =   5805
      Visible         =   0   'False
      Width           =   5685
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   23
         Left            =   1770
         TabIndex        =   10
         ToolTipText     =   "Enter From Order Tc"
         Top             =   810
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPTCFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   24
         Left            =   2445
         TabIndex        =   11
         ToolTipText     =   "Enter From Order Year"
         Top             =   810
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPYYFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   25
         Left            =   2880
         TabIndex        =   12
         ToolTipText     =   "Enter From Order Character"
         Top             =   810
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPCHRFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   26
         Left            =   1770
         TabIndex        =   16
         ToolTipText     =   "Enter From Order Number"
         Top             =   1095
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WSCOPNOFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   27
         Left            =   2700
         TabIndex        =   17
         ToolTipText     =   "Enter From Order Serial"
         Top             =   1095
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCOPSRFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   28
         Left            =   3750
         TabIndex        =   13
         ToolTipText     =   "Enter To Order Tc"
         Top             =   810
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPTCTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   29
         Left            =   4395
         TabIndex        =   14
         ToolTipText     =   "Enter To Order Year"
         Top             =   810
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPYYTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   30
         Left            =   4830
         TabIndex        =   15
         ToolTipText     =   "Enter To Order Character"
         Top             =   810
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOPCHRTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   31
         Left            =   3750
         TabIndex        =   18
         ToolTipText     =   "Enter To Order Number"
         Top             =   1095
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         Mask            =   "######0"
         MaxLength       =   7
         DataType        =   1
         IdName          =   "WSCOPNOTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   32
         Left            =   4680
         TabIndex        =   19
         ToolTipText     =   "Enter To Order Serial"
         Top             =   1095
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         Mask            =   "###0"
         MaxLength       =   4
         DataType        =   1
         IdName          =   "WSCOPSRTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   33
         Left            =   1770
         TabIndex        =   22
         ToolTipText     =   "Enter From Customer"
         Top             =   1665
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOMCMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   34
         Left            =   3750
         TabIndex        =   23
         ToolTipText     =   "Enter To Customer"
         Top             =   1665
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCOMCMCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   35
         Left            =   1770
         TabIndex        =   24
         ToolTipText     =   "Enter From Order Date"
         Top             =   1950
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCOMDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   36
         Left            =   3750
         TabIndex        =   25
         ToolTipText     =   "Enter To Order Date"
         Top             =   1950
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCOMDTTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   37
         Left            =   1770
         TabIndex        =   28
         ToolTipText     =   "Enter From Priority"
         Top             =   2520
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCODPRTCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   38
         Left            =   3750
         TabIndex        =   29
         ToolTipText     =   "Enter To Priority"
         Top             =   2520
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCODPRTCDTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   39
         Left            =   1770
         TabIndex        =   26
         ToolTipText     =   "Enter From Production Delivery Date"
         Top             =   2235
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCODDELDTFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   40
         Left            =   3750
         TabIndex        =   27
         ToolTipText     =   "Enter To Production Delivery Date"
         Top             =   2235
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         IdName          =   "WSCODDELDTTO"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   111
         Left            =   1770
         TabIndex        =   30
         ToolTipText     =   "Sort Records On"
         Top             =   2805
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WSCSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   112
         Left            =   1770
         TabIndex        =   20
         ToolTipText     =   "Enter From Line"
         Top             =   1380
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCOPLINEFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   113
         Left            =   3750
         TabIndex        =   21
         ToolTipText     =   "Enter To Line"
         Top             =   1380
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WSCOPLINETO"
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Line"
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
         Left            =   120
         TabIndex        =   264
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label LblMain 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Scope"
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
         TabIndex        =   263
         Top             =   90
         Width           =   5655
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
         Index           =   72
         Left            =   120
         TabIndex        =   256
         Top             =   2805
         Width           =   885
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
         Index           =   3
         Left            =   3750
         TabIndex        =   168
         Top             =   390
         Width           =   675
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
         Index           =   2
         Left            =   1770
         TabIndex        =   167
         Top             =   390
         Width           =   675
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
         Index           =   10
         Left            =   4590
         TabIndex        =   166
         Top             =   1095
         Width           =   165
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
         Index           =   12
         Left            =   4740
         TabIndex        =   165
         Top             =   810
         Width           =   165
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
         Index           =   13
         Left            =   4320
         TabIndex        =   164
         Top             =   810
         Width           =   105
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Ord No/ Sr"
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
         Left            =   120
         TabIndex        =   163
         Top             =   1095
         Width           =   1185
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
         Index           =   15
         Left            =   2610
         TabIndex        =   162
         Top             =   1095
         Width           =   165
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
         Index           =   17
         Left            =   2790
         TabIndex        =   161
         Top             =   810
         Width           =   165
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   " /"
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
         Left            =   2280
         TabIndex        =   160
         Top             =   810
         Width           =   195
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Ord Tc/ Yy/ Chr"
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
         Left            =   120
         TabIndex        =   159
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   120
         TabIndex        =   158
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Ord Dt"
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
         Left            =   120
         TabIndex        =   157
         Top             =   1950
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Priority"
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
         Left            =   120
         TabIndex        =   156
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Prd Del Dt"
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
         Left            =   120
         TabIndex        =   155
         Top             =   2235
         Width           =   1095
      End
   End
End
Attribute VB_Name = "EmrFrmPpcOrdPln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- General comments on the entry ----
'
'Ppc Entry for Generating Production Plan
'
'---

Option Explicit
'*** For the Toggle Option of Command Buttons
Enum en_OrdPlnFra
  Scop = 0
  ResPln = 1
  UpdStat = 2
  GenPln = 3
  ShwLd = 4
  DsgDet = 5
  RedoPln = 6
  AutoGenPln = 7    '*** (Je 2.14)
End Enum

'*** Arr_Cv stores the causes validation state, Arr_Ed the Ena/Disa state of Cmd Buttons
Dim mArr_Cv As Variant, mArr_Ed As Variant

Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Dim ms_wScOpTcFr As String, ms_wScOpTcTo As String
Dim ms_wScOpYyFr As String, ms_wScOpYyTo As String
Dim ms_wScOpChrFr As String, ms_wScOpChrTo As String
Dim mf_wScOpNoFr As Single, mf_wScOpNoTo As Single
Dim mf_wScOpSrFr As Single, mf_wScOpSrTo As Single
Dim ms_wScOmCmCdFr As String, ms_wScOmCmCdTo As String
Dim ms_wScOmDtFr As String, ms_wScOmDtTo As String
Dim ms_wScOdDelDtFr As String, ms_wScOdDelDtTo As String
Dim ms_wScOdPrtCdFr As String, ms_wScOdPrtCdTo As String
Dim ms_wScOpLineFr As String, ms_wScOpLineTo As String    '*** (Jen 2.14)
Dim ms_wScSrt As String

Dim mf_hMaxLineUtil As Single, mf_hMaxOrdQty As Single, mf_hMinOrdPrcnt As Single
Dim mf_hMinOrdQty As Single

Dim mfArr_AllLineCap(1 To 5) As Single
Dim ms_OrdByFlds As String

'*** Declare variables wPic3DHt, wPic3DWdt, wPicHskHt, wPicHskWdt as single to store the
    'original Height and Width of the 3D and Sketch Picture controls
'*** Declare variables wFraPic3DLft, wFraPic3DTop, wFraPicHskLft, wFraPicHskTop as single
    'to store the original Positions of the 3D and Sketch Picture controls
Dim md_Pic3DHt As Double
Dim md_Pic3DWdt As Double
Dim md_PicHskHt As Double
Dim md_PicHskWdt As Double

Dim mf_FraPic3DHt As Single
Dim mf_FraPic3DWdt As Single
Dim mf_FraPicHskHt As Single
Dim mf_FraPicHskWdt As Single

Dim md_FraPic3DLft As Double
Dim md_FraPic3DTop As Double
Dim md_FraPicHskLft As Double
Dim md_FraPicHskTop As Double

'*** (Jen 3.1.2)
Dim mf_Time As Double   '*** This is required because suddenly the maximum value the 'Interval' property of
                        '    the timer can take is 65535
'*** (Jen 3.1.2)

Private Sub ADC_Load()
  '*** Set the First Key Control property to adc
  '*** Set the First Non Key Control property to the grid GrdPpcCap
  '*** Set the Child Property Of ADC to GrdPpcCap
  '*** Set the Previous and Next Control Properties for the Grid GrdPpcCap
  '*** Set the Hot Key for the Grid GrdPpcCap as 'A'

  'ReDim mArr_Cv(7), mArr_Ed(7)    '*** (Jen 2.14)
  ReDim mArr_Cv(8), mArr_Ed(8)     '*** (Jen 2.14)
  Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06

  Set ADC.FirKeyCtl = ADC
  Set ADC.FirNKeyCtl = GrdOp
  
  ADC.Child = "GrdOp"
  GrdOp.Child = "GrdOpd"
  
  Set GrdOp.PrevCtl = ADC
  Set GrdOp.NextCtl = ADC
  Set GrdOpd.PrevCtl = GrdOp
  Set GrdOpd.NextCtl = GrdOp
  
  GrdOp.HotKey = "A"
  GrdOpd.HotKey = "B"
  
  FraNKeyAll.ZOrder
    
  md_Pic3DHt = Pic3D.Height
  md_Pic3DWdt = Pic3D.Width
  md_PicHskHt = PicHsk.Height
  md_PicHskWdt = PicHsk.Width
  
  mf_FraPic3DHt = FraPic3D.Height
  mf_FraPic3DWdt = FraPic3D.Width
  mf_FraPicHskHt = FraPicHsk.Height
  mf_FraPicHskWdt = FraPicHsk.Width
  
  md_FraPic3DLft = FraPic3D.Left
  md_FraPic3DTop = FraPic3D.Top
  md_FraPicHskLft = FraPicHsk.Left
  md_FraPicHskTop = FraPicHsk.Top


  '*** (Jen 2.14 Next)
  Call Form_Activate
  Call GrdPropSetPPC(GrdOpd, "Opd")
  Call GrdPropSetPPC(GrdOpd, "wOpdLoad")
  Call GrdPropSetPPC(GrdSL, "wSl")
  '*** (Jen 2.14 Next)
  
  mf_Time = 0     '*** (Jen 3.1.2)
End Sub
Private Sub ADC_KeyWhen()

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(gs_CoCd, "Plan Master")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

  If ms_wScOpTcFr <> "" Then ADC("wScOpTcFr") = ms_wScOpTcFr
  If ms_wScOpTcTo <> "" Then ADC("wScOpTcTo") = ms_wScOpTcTo
  If ms_wScOpYyFr <> "" Then ADC("wScOpYyFr") = ms_wScOpYyFr
  If ms_wScOpYyTo <> "" Then ADC("wScOpYyTo") = ms_wScOpYyTo
  If ms_wScOpChrFr <> "" Then ADC("wScOpChrFr") = ms_wScOpChrFr
  If ms_wScOpChrTo <> "" Then ADC("wScOpChrTo") = ms_wScOpChrTo
  If mf_wScOpNoFr <> 0 Then ADC("wScOpNoFr") = mf_wScOpNoFr
  If mf_wScOpNoTo <> 0 Then ADC("wScOpNoTo") = mf_wScOpNoTo
  If mf_wScOpSrFr <> 0 Then ADC("wScOpSrFr") = mf_wScOpSrFr
  If mf_wScOpSrTo <> 0 Then ADC("wScOpSrTo") = mf_wScOpSrTo
  If ms_wScOmCmCdFr <> "" Then ADC("wScOmCmCdFr") = ms_wScOmCmCdFr
  If ms_wScOmCmCdTo <> "" Then ADC("wScOmCmCdTo") = ms_wScOmCmCdTo
  If ms_wScOmDtFr <> "" Then ADC("wScOmDtFr") = ms_wScOmDtFr
  If ms_wScOmDtTo <> "" Then ADC("wScOmDtTo") = ms_wScOmDtTo
  If ms_wScOdDelDtFr <> "" Then ADC("wScOdDelDtFr") = ms_wScOdDelDtFr
  If ms_wScOdDelDtTo <> "" Then ADC("wScOdDelDtTo") = ms_wScOdDelDtTo
  If ms_wScOdPrtCdFr <> "" Then ADC("wScOdPrtCdFr") = ms_wScOdPrtCdFr
  If ms_wScOdPrtCdTo <> "" Then ADC("wScOdPrtCdTo") = ms_wScOdPrtCdTo
  '*** (Jen 2.14)
  If ms_wScOpLineFr <> "" Then ADC("wScOpLineFr") = ms_wScOpLineFr
  If ms_wScOpLineTo <> "" Then ADC("wScOpLineTo") = ms_wScOpLineTo
  '*** (Jen 2.14)
  If ms_wScSrt <> "" Then ADC("wScSrt") = ms_wScSrt
  
  Call EnaDisaCmds(True)
  CmdSC.Enabled = True
  Call HideAllFras
End Sub
Private Sub ADC_KeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  If Not moCn.RecSeek("Select 'x' from Head where HCoCd= '" + gs_CoCd + "' and " + _
         "HCd= '" + ctSelfCmCd + "' and hPpcYN= 'Y'") Then
    pr_Cancel = True: pr_ErrMsg = "Cannot Edit This Entry As Permission For PPC Module Not Given": Exit Sub
  End If
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Hide the Scope frame if it is still visible when the User Clicks on Add/Find
  If FraSC.Visible = True Then
    Call DispFra(Scop)
  End If
  
  Call EnaDisaCmds(False)
  CmdSC.Enabled = False
  CmdResPln.Enabled = True
  CmdUpdStat.Enabled = True
  CmdGenPln.Enabled = True
  CmdShwLd.Enabled = True
  CmdDsgDet.Enabled = True
  CmdRedoPln.Enabled = True
  CmdAutoGenPln.Enabled = True     '*** (Jen 2.14)

  ms_wScOpTcFr = ADC("wScOpTcFr"): ms_wScOpTcTo = ADC("wScOpTcTo")
  ms_wScOpYyFr = ADC("wScOpYyFr"): ms_wScOpYyTo = ADC("wScOpYyTo")
  ms_wScOpChrFr = ADC("wScOpChrFr"): ms_wScOpChrTo = ADC("wScOpChrTo")
  mf_wScOpNoFr = ADC("wScOpNoFr"): mf_wScOpNoTo = ADC("wScOpNoTo")
  mf_wScOpSrFr = ADC("wScOpSrFr"): mf_wScOpSrTo = ADC("wScOpSrTo")
  ms_wScOmCmCdFr = ADC("wScOmCmCdFr"): ms_wScOmCmCdTo = ADC("wScOmCmCdTo")
  ms_wScOmDtFr = ADC("wScOmDtFr"): ms_wScOmDtTo = ADC("wScOmDtTo")
  ms_wScOdDelDtFr = ADC("wScOdDelDtFr"): ms_wScOdDelDtTo = ADC("wScOdDelDtTo")
  ms_wScOdPrtCdFr = ADC("wScOdPrtCdFr"): ms_wScOdPrtCdTo = ADC("wScOdPrtCdTo")
  ms_wScOpLineFr = ADC("wScOpLineFr"): ms_wScOpLineTo = ADC("wScOpLineTo")    '*** (Jen 2.14)
  ms_wScSrt = ADC("wScSrt")

  ADC("wScOpTcFr") = "": ADC("wScOpTcTo") = ""
  ADC("wScOpYyFr") = "": ADC("wScOpYyTo") = ""
  ADC("wScOpChrFr") = "": ADC("wScOpChrTo") = ""
  ADC("wScOpNoFr") = "": ADC("wScOpNoTo") = ""
  ADC("wScOpSrFr") = "": ADC("wScOpSrTo") = ""
  ADC("wScOmCmCdFr") = "": ADC("wScOmCmCdTo") = ""
  ADC("wScOmDtFr") = "": ADC("wScOmDtTo") = ""
  ADC("wScOdDelDtFr") = "": ADC("wScOdDelDtTo") = ""
  ADC("wScOdPrtCdFr") = "": ADC("wScOdPrtCdTo") = ""
  ADC("wScOpLineFr") = "": ADC("wScOpLineTo") = ""    '*** (Jen 2.14)
  
  mf_hMaxLineUtil = Round(moCn.GetFldVal("Select hMaxLineUtil from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' "), 3)
  mf_hMaxOrdQty = Round(moCn.GetFldVal("Select hMaxOrdQty from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' "), 1)
  mf_hMinOrdPrcnt = Round(moCn.GetFldVal("Select hMinOrdPrcnt from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' "), 3)
  mf_hMinOrdQty = Round(moCn.GetFldVal("Select hMinOrdQty from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' "), 1)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  
  Select Case UCase(IdName)
  Case Is = UCase("wScOpTcFr"), UCase("wScOpTcTo"), UCase("wUsOpTcFr"), UCase("wUsOpTcTo"), _
            UCase("wGpOpTcFr"), UCase("wGpOpTcTo"), UCase("wRpOpTcFr"), UCase("wRpOpTcTo"), _
            UCase("wAgOpTcFr"), UCase("wAgOpTcTo")
    '****** (Jen 2.14) added fields wAgOpTcFr, wAgOpTcTo
    Call HlpList.PMCd("TC", "'SO'")
  Case Is = UCase("wScOpYyFr"), UCase("wScOpYyTo"), UCase("wUsOpYyFr"), UCase("wUsOpYyTo"), _
            UCase("wGpOpYyFr"), UCase("wGpOpYyTo"), UCase("wRpOpYyFr"), UCase("wRpOpYyTo"), _
            UCase("wAgOpYyFr"), UCase("wAgOpYyTo")
    '****** (Jen 2.14) added fields wAgOpYyFr, wAgOpYyTo
    Call HlpList.PMCd("YY")
  
  Case Is = UCase("wScOpChrFr")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wScOpTcFr"))
  Case Is = UCase("wScOpChrTo")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wScOpTcTo"))
  Case Is = UCase("wUsOpChrFr")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wUsOpTcFr"))
  Case Is = UCase("wUsOpChrTo")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wUsOpTcTo"))
  Case Is = UCase("wGpOpChrFr")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wGpOpTcFr"))
  Case Is = UCase("wGpOpChrTo")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wGpOpTcTo"))
  Case Is = UCase("wRpOpChrFr")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wRpOpTcFr"))
  Case Is = UCase("wRpOpChrTo")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wRpOpTcTo"))
  
  '*** (Jen 2.14)
  Case Is = UCase("wAgOpChrFr")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wAgOpTcFr"))
  Case Is = UCase("wAgOpChrTo")
     Call HlpList.vPSCd(gs_CoCd, "CHR", ADC("wAgOpTcTo"))
  '*** (Jen 2.14)
  
  Case Is = UCase("wScOpNoFr")
     Call HlpList.OmNo(gs_CoCd, ADC("wScOpTcFr"), ADC("wScOpYyFr"), ADC("wScOpChrFr"))
  Case Is = UCase("wScOpNoTo")
     Call HlpList.OmNo(gs_CoCd, ADC("wScOpTcTo"), ADC("wScOpYyTo"), ADC("wScOpChrTo"))
  Case Is = UCase("wUsOpNoFr")
     Call HlpList.OmNo(gs_CoCd, ADC("wUsOpTcFr"), ADC("wUsOpYyFr"), ADC("wUsOpChrFr"))
  Case Is = UCase("wUsOpNoTo")
     Call HlpList.OmNo(gs_CoCd, ADC("wUsOpTcTo"), ADC("wUsOpYyTo"), ADC("wUsOpChrTo"))
  Case Is = UCase("wGpOpNoFr")
     Call HlpList.OmNo(gs_CoCd, ADC("wGpOpTcFr"), ADC("wGpOpYyFr"), ADC("wGpOpChrFr"))
  Case Is = UCase("wGpOpNoTo")
     Call HlpList.OmNo(gs_CoCd, ADC("wGpOpTcTo"), ADC("wGpOpYyTo"), ADC("wGpOpChrTo"))
  Case Is = UCase("wRpOpNoFr")
     Call HlpList.OmNo(gs_CoCd, ADC("wRpOpTcFr"), ADC("wRpOpYyFr"), ADC("wRpOpChrFr"))
  Case Is = UCase("wROpNoTo")
     Call HlpList.OmNo(gs_CoCd, ADC("wRpOpTcTo"), ADC("wRpOpYyTo"), ADC("wRpOpChrTo"))
  
  '*** (Jen 2.14)
  Case Is = UCase("wAgOpNoFr")
     Call HlpList.OmNo(gs_CoCd, ADC("wAgOpTcFr"), ADC("wAgOpYyFr"), ADC("wAgOpChrFr"))
  Case Is = UCase("wAgOpNoTo")
     Call HlpList.OmNo(gs_CoCd, ADC("wAgOpTcTo"), ADC("wAgOpYyTo"), ADC("wAgOpChrTo"))
  '*** (Jen 2.14)
  
  Case Is = UCase("wScOpSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wScOpTcFr"), ADC("wScOpYyFr"), ADC("wScOpChrFr"), ADC("wScOpNoFr"))
  Case Is = UCase("wScOpSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wScOpTcTo"), ADC("wScOpYyTo"), ADC("wScOpChrTo"), ADC("wScOpNoTo"))
  Case Is = UCase("wUsOpSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wUsOpTcFr"), ADC("wUsOpYyFr"), ADC("wUsOpChrFr"), ADC("wUsOpNoFr"))
  Case Is = UCase("wUsOpSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wUsOpTcTo"), ADC("wUsOpYyTo"), ADC("wUsOpChrTo"), ADC("wUsOpNoTo"))
  Case Is = UCase("wGpOpSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wGpOpTcFr"), ADC("wGpOpYyFr"), ADC("wGpOpChrFr"), ADC("wGpOpNoFr"))
  Case Is = UCase("wGpOpSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wGpOpTcTo"), ADC("wGpOpYyTo"), ADC("wGpOpChrTo"), ADC("wGpOpNoTo"))
  Case Is = UCase("wRpOpSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wRpOpTcFr"), ADC("wRpOpYyFr"), ADC("wRpOpChrFr"), ADC("wRpOpNoFr"))
  Case Is = UCase("wROpSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wRpOpTcTo"), ADC("wRpOpYyTo"), ADC("wRpOpChrTo"), ADC("wRpOpNoTo"))
  
  '*** (Jen 2.14)
  Case Is = UCase("wAgOpSrFr")
    Call HlpList.OdSr(gs_CoCd, ADC("wAgOpTcFr"), ADC("wAgOpYyFr"), ADC("wAgOpChrFr"), ADC("wAgOpNoFr"))
  Case Is = UCase("wAgOpSrTo")
    Call HlpList.OdSr(gs_CoCd, ADC("wAgOpTcTo"), ADC("wAgOpYyTo"), ADC("wAgOpChrTo"), ADC("wAgOpNoTo"))
  '*** (Jen 2.14)
  
  Case Is = UCase("wScOmCmCdFr"), UCase("wScOmCmCdTo"), UCase("wUsOmCmCdFr"), UCase("wUsOmCmCdTo"), _
            UCase("wGpOmCmCdFr"), UCase("wGpOmCmCdTo"), UCase("wRpOmCmCdFr"), UCase("wRpOmCmCdTo"), _
            UCase("wAgOmCmCdFr"), UCase("wAgOmCmCdTo")
    '****** (Jen 2.14) added fields wAgOmCmCdFr, wAgOmCmCdTo
    Call HlpList.CustCd("C")
  Case Is = UCase("wScOdPrtCdFr"), UCase("wScOdPrtCdTo"), UCase("wUsOdPrtCdFr"), UCase("wUsOdPrtCdTo"), _
            UCase("wGpOdPrtCdFr"), UCase("wGpOdPrtCdTo"), UCase("wRpOdPrtCdFr"), UCase("wRpOdPrtCdTo"), _
            UCase("wAgOdPrtCdFr"), UCase("wAgOdPrtCdTo")
    '****** (Jen 2.14) added fields wAgOdPrtCdFr, wAgOdPrtCdTo
    Call HlpList.PMCd("PRTCD")
  
  Case Is = UCase("wSlLine")
    Call HlpList.vPMCd(gs_CoCd, "LINE")
  Case Is = UCase("wScSrt")
    Call HlpList.PMCd("SRTOP")
  
  '*** (Jen 2.14)
  Case Is = UCase("wScOpLineFr"), UCase("wScOpLineTo"), UCase("wUsOpLineFr"), UCase("wUsOpLineTo"), _
            UCase("wGpOpLineFr"), UCase("wGpOpLineTo"), UCase("wRpOpLineFr"), UCase("wRpOpLineTo"), _
            UCase("wAgOpLineFr"), UCase("wAgOpLineTo")
    Call HlpList.vPMCd(gs_CoCd, "Line")
  Case Is = UCase("wUsGetNewOrd"), UCase("wAgGetNewOrd")
    Call HlpList.PMCd("YN")
  '*** (Jen 2.14)
  End Select
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("wOpDldP1")
    If ADC.Mode = xNorm And (GrdOp.Mode = fgmnorm) Then
      ADC("wOpDldP1") = GrdOp.Value(RowNum, "wP1LdPpc")
    
      'ADC("wOpDldWSet") = moCn.GetFldVal("Select IsNull" + _
                          "((Select DldwSet from DsgLoad " + _
                            "where DldDmCd= '" + GrdOp.Value(RowNum, "wOdDmCd") + "' " + _
                            "and DldWH= '" + GrdOp.Value(RowNum, "wOdWH") + "') " + _
                          ", " + _
                          "IsNull((Select DldwSet from DsgLoad " + _
                                 "where DldDmCd= '' and DldWH= '" + GrdOp.Value(RowNum, "wOdWH") + "' " + _
                                 "and DldPrdCtg= '" + moCn.GetFldVal("Select DmPrdCtg from DsgMst " + _
                                 "where DmTcTyp='DM' and DmCd= '" + GrdOp.Value(RowNum, "wOdDmCd") + "' " + _
                                 "and DmSz= ''") + "'), " + _
                                 "0) " + _
                           ")")
    End If
  Case Is = UCase("wOpDldP2")
    If ADC.Mode = xNorm And (GrdOp.Mode = fgmnorm) Then
      ADC("wOpDldP2") = GrdOp.Value(RowNum, "wP2LdPpc")
    End If
  Case Is = UCase("wOpDldP3")
    If ADC.Mode = xNorm And (GrdOp.Mode = fgmnorm) Then
      ADC("wOpDldP3") = GrdOp.Value(RowNum, "wP3LdPpc")
    End If
  Case Is = UCase("wOpDldP4")
    If ADC.Mode = xNorm And (GrdOp.Mode = fgmnorm) Then
      ADC("wOpDldP4") = GrdOp.Value(RowNum, "wP4LdPpc")
    End If
  Case Is = UCase("wOpDldP5")
    If ADC.Mode = xNorm And (GrdOp.Mode = fgmnorm) Then
      ADC("wOpDldP5") = GrdOp.Value(RowNum, "wP5LdPpc")
    End If
  End Select
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("wScOpTcFr")
    If ADC("wScOpTcTo") = "" Then ADC("wScOpTcTo") = pv_NewValue
  Case Is = UCase("wScOpYyFr")
    If ADC("wScOpYyTo") = "" Then ADC("wScOpYyTo") = pv_NewValue
  Case Is = UCase("wScOpChrFr")
    If ADC("wScOpChrTo") = "" Then ADC("wScOpChrTo") = pv_NewValue
  Case Is = UCase("wScOpNoFr")
    If ADC("wScOpNoTo") = 0 Then ADC("wScOpNoTo") = pv_NewValue
  Case Is = UCase("wScOpSrFr")
    If ADC("wScOpSrTo") = 0 Then ADC("wScOpSrTo") = pv_NewValue
  Case Is = UCase("wScOmCmCdFr")
    If ADC("wScOmCmCdTo") = "" Then ADC("wScOmCmCdTo") = pv_NewValue
  Case Is = UCase("wScOmDtFr")
    If ADC("wScOmDtTo") = "" Then ADC("wScOmDtTo") = pv_NewValue
  Case Is = UCase("wScOdDelDtFr")
    If ADC("wScOdDelDtTo") = "" Then ADC("wScOdDelDtTo") = pv_NewValue
  Case Is = UCase("wScOdPrtCdFr")
    If ADC("wScOdPrtCdTo") = "" Then ADC("wScOdPrtCdTo") = pv_NewValue
  
  '*** (Jen 2.14)
  Case Is = UCase("wScOpLineFr")
    If ADC("wScOpLineTo") = "" Then ADC("wScOpLineTo") = pv_NewValue
  '*** (Jen 2.14)
  
  Case Is = UCase("wSlLine")
    Cancel = Not moCn.RecSeek("select vPMCd from vParam where vPCoCd= '" + gs_CoCd + "' and " + _
             "vPTyp= 'LINE' and vPMCd = '" + pv_NewValue + "' and vPSCd= ''")
    If Cancel = True Then ErrMsg = "Invalid Line": Exit Sub
  Case Is = UCase("wScSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTOP' and PMCd= '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    End If
  
  '*** (Jen 2.14)
  Case Is = UCase("wUsGetNewOrd"), UCase("wAgGetNewOrd")
    Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'YN' and PMCd= '" + pv_NewValue + "'"))
    ErrMsg = "Invalid 'Get New Orders' Option": Exit Sub
  '*** (Jen 2.14)
  
  End Select
End Sub
Private Sub GrdOp_RowWhen(ByVal RowNum As Integer)
  With GrdOp
    ADC("wOpWipP1") = .Value(RowNum, "OpWipP1")
    ADC("wOpWipP2") = .Value(RowNum, "OpWipP2")
    ADC("wOpWipP3") = .Value(RowNum, "OpWipP3")
    ADC("wOpWipP4") = .Value(RowNum, "OpWipP4")
    ADC("wOpWipP5") = .Value(RowNum, "OpWipP5")
    
    ADC("wOpDldP1") = .Value(RowNum, "wP1LdPpc")
    ADC("wOpDldP2") = .Value(RowNum, "wP2LdPpc")
    ADC("wOpDldP3") = .Value(RowNum, "wP3LdPpc")
    ADC("wOpDldP4") = .Value(RowNum, "wP4LdPpc")
    ADC("wOpDldP5") = .Value(RowNum, "wP5LdPpc")
  End With
End Sub
Private Sub GrdOp_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  With GrdOp
    Select Case UCase(ColName)
    Case Is = UCase("OpTc")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Already Exists": Exit Sub
      Call HlpList.PMCd("TC")
    Case Is = UCase("OpYy")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Already Exists": Exit Sub
      Call HlpList.PMCd("YY")
    Case Is = UCase("OpChr")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Already Exists": Exit Sub
      Call HlpList.vPSCd(gs_CoCd, "CHR", .Value(RowNum, "OpTc"))
    Case Is = UCase("OpNo")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Already Exists": Exit Sub
      Call HlpList.OmNo(gs_CoCd, .Value(RowNum, "OpTc"), .Value(RowNum, "OpYy"), .Value(RowNum, "OpChr"))
    Case Is = UCase("OpSr")
      If Not .IsNew(RowNum) Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Already Exists": Exit Sub
      Call HlpList.OdSr(gs_CoCd, .Value(RowNum, "OpTc"), .Value(RowNum, "OpYy"), .Value(RowNum, "OpChr"), .Value(RowNum, "OpNo"))
    Case Is = UCase("OpLine")
      If GrdOpd.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Exists In Child Grid": Exit Sub
      Call HlpList.vPMCd(gs_CoCd, "LINE")
    Case Is = UCase("OpWDiaAvlblDt"), UCase("OpHDiaAvlblDt")
      If GrdOpd.Rows - 1 > 0 Then Cancel = True: ErrMsg = "Cannot Edit This Value As Rec Exists In Child Grid": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdOp_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdOp
    Select Case UCase(ColName)
    Case Is = UCase("wOdDelDt")
      '*** (Bef 2.14)
      'If .Mode = fgmnorm Or .Mode = fgminit Then
      '  .Value(RowNum, "wOdDelDt") = moCn.GetFldVal("Select OdDelDt from OrdDsg where " + _
      '         "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
      '         "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
      '         "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
      '         "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
      '         "OdSr= " + CStr(.Value(RowNum, "OpSr")))
      'End If
      '*** (Bef 2.14)
      
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
        Dim wRs_DsgDet As MwfLib.MDORowSet
        Set wRs_DsgDet = moCn.OpenRes("Select OdDelDt, OdDmCd, OdWH, OdPrtCd, OdPrdQty, (OdPrdQty- OdFgQty) as qBal " + _
                         "from OrdDsg where OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' " + _
                         "and OdYy= '" + .Value(RowNum, "OpYy") + "' and OdChr= '" + .Value(RowNum, "OpChr") + "' " + _
                         "and OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and OdSr= " + CStr(.Value(RowNum, "OpSr")))
        If Not (wRs_DsgDet.EOF Or wRs_DsgDet.BOF) Then
          .Value(RowNum, "wOdDelDt") = wRs_DsgDet!OdDelDt
          .Value(RowNum, "wOdDmCd") = wRs_DsgDet!OdDmCd
          .Value(RowNum, "wOdWH") = wRs_DsgDet!OdWH
          .Value(RowNum, "wOdPrtCd") = wRs_DsgDet!OdPrtCd
          .Value(RowNum, "wOdPrdQty") = wRs_DsgDet!OdPrdQty
          .Value(RowNum, "wBalQty") = wRs_DsgDet!qBal
        End If
        Set wRs_DsgDet = Nothing
      End If
      '*** (Jen 2.14)
      
    Case Is = UCase("OpLine")
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpLine") = moCn.GetFldVal("Select OdLine from OrdDsg where " + _
               "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
               "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
               "OdSr= " + CStr(.Value(RowNum, "OpSr")))
      End If
    Case Is = UCase("OpWDiaAvlblDt")
      If GrdOp.Mode = fgmnorm Then
        .Value(RowNum, "OpWDiaAvlblDt") = moCn.GetFldVal("Select OdwDiaAvlblDt from OrdDsg where " + _
               "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
               "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
               "OdSr= " + CStr(.Value(RowNum, "OpSr")))
      End If
    Case Is = UCase("OpHDiaAvlblDt")
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpHDiaAvlblDt") = moCn.GetFldVal("Select OdhDiaAvlblDt from OrdDsg where " + _
               "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
               "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
               "OdSr= " + CStr(.Value(RowNum, "OpSr")))
      End If
        
    '*** (Bef 2.14)
    'Case Is = UCase("wOdDmCd")
    '  '*** recalc on GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    .Value(RowNum, "wOdDmCd") = moCn.GetFldVal("Select OdDmCd from OrdDsg where " + _
    '           "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
    '           "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
    '           "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
    '           "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
    '           "OdSr= " + CStr(.Value(RowNum, "OpSr")))
    '  End If
    'Case Is = UCase("wOdWH")
    '  '*** recalc on GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    .Value(RowNum, "wOdWH") = moCn.GetFldVal("Select OdWH from OrdDsg where " + _
    '           "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
    '           "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
    '           "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
    '           "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
    '           "OdSr= " + CStr(.Value(RowNum, "OpSr")))
    '  End If
    'Case Is = UCase("wOdPrtCd")
    '  '*** recalc on GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    .Value(RowNum, "wOdPrtCd") = moCn.GetFldVal("Select OdPrtCd from OrdDsg where " + _
    '           "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
    '           "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
    '           "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
    '           "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
    '           "OdSr= " + CStr(.Value(RowNum, "OpSr")))
    '  End If
    'Case Is = UCase("wOdPrdQty")
    '  '*** recalc on GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    .Value(RowNum, "wOdPrdQty") = moCn.GetFldVal("Select OdPrdQty from OrdDsg where " + _
    '           "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
    '           "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
    '           "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
    '           "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
    '           "OdSr= " + CStr(.Value(RowNum, "OpSr")))
    '  End If
    '*** (Bef 2.14)
    
    Case Is = UCase("OpWipP1")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpWipP1") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " and " + _
               "BPrdStg= 'P1' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
    Case Is = UCase("OpWipP2")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpWipP2") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " and " + _
               "BPrdStg= 'P2' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
    Case Is = UCase("OpWipP3")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpWipP3") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " and " + _
               "BPrdStg= 'P3' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
    Case Is = UCase("OpWipP4")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpWipP4") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " and " + _
               "BPrdStg= 'P4' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
    Case Is = UCase("OpWipP5")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Then
        .Value(RowNum, "OpWipP5") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " and " + _
               "BPrdStg= 'P5' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
    Case Is = UCase("wFlrQty")
      '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
      If .Mode = fgmnorm Or .Mode = fgminit Then
        .Value(RowNum, "wFlrQty") = moCn.GetFldVal("Select sum(BQty) from Bag where " + _
               "BCoCd= '" + gs_CoCd + "' and BOdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
               "BOdYy= '" + .Value(RowNum, "OpYy") + "' and BOdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
               "BOdNo= " + CStr(.Value(RowNum, "OpNo")) + " and BOdSr= " + CStr(.Value(RowNum, "OpSr")) + " " + _
               "and BCls= 'N' and BLoc<> 'PFG' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') ")
               
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
      End If
      
    '*** (Bef 2.14)
    'Case Is = UCase("wBalQty")
    '  '*** recalc on GrdOp(OPTC),GrdOp(OPYY),GrdOp(OPCHR),GrdOp(OPNO),GrdOp(OPSR)
    '  If .Mode = fgmnorm Or .Mode = fgminit Then
    '    .Value(RowNum, "wBalQty") = moCn.GetFldVal("Select OdPrdQty- OdFgQty from OrdDsg where " + _
    '           "OdCoCd= '" + gs_CoCd + "' and OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
    '           "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
    '           "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
    '           "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
    '           "OdSr= " + CStr(.Value(RowNum, "OpSr")))
    '  End If
    '*** (Bef 2.14)
    Case Is = UCase("wP1LdPpc")
      If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then
        Dim wRs_DsgLdPerPc As MwfLib.MDORowSet
        Set wRs_DsgLdPerPc = moCn.OpenRes("Select IsNull(DL1.DldP1, IsNull(DL2.DldP1, 0)) as qDldP1, " + _
                             "IsNull(DL1.DldP2, IsNull(DL2.DldP2, 0)) as qDldP2, " + _
                             "IsNull(DL1.DldP3, IsNull(DL2.DldP3, 0)) as qDldP3, " + _
                             "IsNull(DL1.DldP4, IsNull(DL2.DldP4, 0)) as qDldP4, " + _
                             "IsNull(DL1.DldP5, IsNull(DL2.DldP5, 0)) as qDldP5 from OrdDsg " + _
                             "join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
                             "left outer join DsgLoad DL1 on DL1.DldDmCd= OdDmCd and DL1.DldWH= OdWH " + _
                             "left outer join DsgLoad DL2 on DL2.DldPrdCtg= DmPrdCtg and " + _
                             "DL2.DldDmCd= '' and DL2.DldWH= OdWH where OdCoCd= '" + gs_CoCd + "' and " + _
                             "OdTc= '" + .Value(RowNum, "OpTc") + "' and " + _
                             "OdYy= '" + .Value(RowNum, "OpYy") + "' and " + _
                             "OdChr= '" + .Value(RowNum, "OpChr") + "' and " + _
                             "OdNo= " + CStr(.Value(RowNum, "OpNo")) + " and " + _
                             "OdSr= " + CStr(.Value(RowNum, "OpSr")))
        
        If Not (wRs_DsgLdPerPc.EOF Or wRs_DsgLdPerPc.BOF) Then
          .Value(RowNum, "wP1LdPpc") = wRs_DsgLdPerPc!qDldP1
          .Value(RowNum, "wP2LdPpc") = wRs_DsgLdPerPc!qDldP2
          .Value(RowNum, "wP3LdPpc") = wRs_DsgLdPerPc!qDldP3
          .Value(RowNum, "wP4LdPpc") = wRs_DsgLdPerPc!qDldP4
          .Value(RowNum, "wP5LdPpc") = wRs_DsgLdPerPc!qDldP5
        End If
        Set wRs_DsgLdPerPc = Nothing
      End If
    End Select
  End With
End Sub
Private Sub GrdOp_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(ColName)
  Case Is = UCase("OpWDiaAvlblDt")
    If NewValue = "" Or NewValue = "01/01/80" Then _
      Cancel = True: ErrMsg = "Invalid Diamond Availability Date For Wax Set": Exit Sub
  Case Is = UCase("OpHDiaAvlblDt")
    If NewValue = "" Or NewValue = "01/01/80" Then _
      Cancel = True: ErrMsg = "Invalid Diamond Availability Date For Hand Set": Exit Sub
  Case Is = UCase("OpLine")
    Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select vPMCd From vParam Where " + _
             "vPCoCd= '" + gs_CoCd + "' and vPTyp= 'LINE' And vPMCd = '" + NewValue + "'"))
    If Cancel = True Then ErrMsg = "Invalid Line No": Exit Sub
  End Select
End Sub
Private Sub GrdOp_InitKey(ByVal RowNum As Integer)
  With GrdOp
    .Store "OpCoCd", gs_CoCd
  End With
End Sub
Private Sub GrdOpd_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdOpd
    Select Case UCase(ColName)
    Case Is = UCase("wOpdLoadP1")
      'If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then    '*** (Bef 2.14)
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
      '*** (Jen 2.14)
        .Value(RowNum, "wOpdLoadP1") = .Value(RowNum, "OpdP1") * GrdOp.Value(GrdOp.Row, "wP1LdPpc")      ' ADC("wOpDldWSet")
      End If
    Case Is = UCase("wOpdLoadP2")
      'If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then    '*** (Bef 2.14)
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
      '*** (Jen 2.14)
        .Value(RowNum, "wOpdLoadP2") = .Value(RowNum, "OpdP2") * GrdOp.Value(GrdOp.Row, "wP2LdPpc")
      End If
    Case Is = UCase("wOpdLoadP3")
      'If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then    '*** (Bef 2.14)
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
      '*** (Jen 2.14)
        .Value(RowNum, "wOpdLoadP3") = .Value(RowNum, "OpdP3") * GrdOp.Value(GrdOp.Row, "wP3LdPpc")
      End If
    Case Is = UCase("wOpdLoadP4")
      'If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then    '*** (Bef 2.14)
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
      '*** (Jen 2.14)
        .Value(RowNum, "wOpdLoadP4") = .Value(RowNum, "OpdP4") * GrdOp.Value(GrdOp.Row, "wP4LdPpc")
      End If
    Case Is = UCase("wOpdLoadP5")
      'If .Mode = fgmnorm Or .Mode = fgminit Or .Mode = fgmCopy Then    '*** (Bef 2.14)
      '*** (Jen 2.14)
      If .Mode = fgmnorm Or .Mode = fgminit Then
      '*** (Jen 2.14)
        .Value(RowNum, "wOpdLoadP5") = .Value(RowNum, "OpdP5") * GrdOp.Value(GrdOp.Row, "wP5LdPpc")
      End If
    End Select
  End With
End Sub
Private Sub GrdOpd_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(ColName)
  Case Is = UCase("OpdPlnDt")
    If NewValue = "" Or NewValue = "01/01/80" Then Cancel = True: ErrMsg = "Invalid Plan Date": Exit Sub
  Case Is = UCase("OpdP1"), UCase("OpdP2"), UCase("OpdP3"), UCase("OpdP4"), UCase("OpdP5")
    If NewValue > GrdOp.Value(GrdOp.Row, "wBalQty") Then Cancel = True: ErrMsg = "Plan Qty Cannot Be More Than Balance Qty": Exit Sub
  End Select
End Sub
Private Sub GrdOpd_InitKey(ByVal RowNum As Integer)
  With GrdOpd
    .Store "OpdCoCd", gs_CoCd
    .Store "OpdTc", GrdOp.Value(GrdOp.Row, "OpTc")
    .Store "OpdYy", GrdOp.Value(GrdOp.Row, "OpYy")
    .Store "OpdChr", GrdOp.Value(GrdOp.Row, "OpChr")
    .Store "OpdNo", GrdOp.Value(GrdOp.Row, "OpNo")
    .Store "OpdSr", GrdOp.Value(GrdOp.Row, "OpSr")
  End With
End Sub
Private Sub GrdSL_RowWhen(ByVal RowNum As Integer)
  Dim wiArr_LineCap(1 To 5) As Single, wiArr_CapAll(1 To 5) As Single
  Dim wRs_PpcCap As MDORowSet
  Set wRs_PpcCap = moCn.OpenRes("Select * from PpcCap where PcCoCd= '" + gs_CoCd + "' and " + _
                   "PcLine in ('ZZZ', '" + ADC("wSlLine") + "') ")
  
  Do While Not (wRs_PpcCap.EOF Or wRs_PpcCap.BOF)
    If UCase(wRs_PpcCap!PcLine) = "ZZZ" Then
      wiArr_CapAll(1) = wRs_PpcCap!PcP1
      wiArr_CapAll(2) = wRs_PpcCap!PcP2
      wiArr_CapAll(3) = wRs_PpcCap!PcP3
      wiArr_CapAll(4) = wRs_PpcCap!PcP4
      wiArr_CapAll(5) = wRs_PpcCap!PcP5
    Else
      wiArr_LineCap(1) = wRs_PpcCap!PcP1
      wiArr_LineCap(2) = wRs_PpcCap!PcP2
      wiArr_LineCap(3) = wRs_PpcCap!PcP3
      wiArr_LineCap(4) = wRs_PpcCap!PcP4
      wiArr_LineCap(5) = wRs_PpcCap!PcP5
    End If
    wRs_PpcCap.MoveNext
  Loop
  
  With GrdSL
    ADC("wSlLineLdP1") = .Value(RowNum, "wSlP1Ld")
    ADC("wSlLineLdP2") = .Value(RowNum, "wSlP2Ld")
    ADC("wSlLineLdP3") = .Value(RowNum, "wSlP3Ld")
    ADC("wSlLineLdP4") = .Value(RowNum, "wSlP4Ld")
    ADC("wSlLineLdP5") = .Value(RowNum, "wSlP5Ld")
  
    ADC("wSlBalCapP1") = wiArr_LineCap(1) - .Value(RowNum, "wSlP1Ld")
    ADC("wSlBalCapP2") = wiArr_LineCap(2) - .Value(RowNum, "wSlP2Ld")
    ADC("wSlBalCapP3") = wiArr_LineCap(3) - .Value(RowNum, "wSlP3Ld")
    ADC("wSlBalCapP4") = wiArr_LineCap(4) - .Value(RowNum, "wSlP4Ld")
    ADC("wSlBalCapP5") = wiArr_LineCap(5) - .Value(RowNum, "wSlP5Ld")
  
    ADC("wSlBalCapAllP1") = wiArr_CapAll(1) - .Value(RowNum, "wSlP1LdAll")
    ADC("wSlBalCapAllP2") = wiArr_CapAll(2) - .Value(RowNum, "wSlP2LdAll")
    ADC("wSlBalCapAllP3") = wiArr_CapAll(3) - .Value(RowNum, "wSlP3LdAll")
    ADC("wSlBalCapAllP4") = wiArr_CapAll(4) - .Value(RowNum, "wSlP4LdAll")
    ADC("wSlBalCapAllP5") = wiArr_CapAll(5) - .Value(RowNum, "wSlP5LdAll")
  End With
End Sub
Private Sub GrdSL_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** No change allowed in the fields of the working grid GrdSL of the Show Cap/Load frame
  Cancel = True
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'AB' Entry, there is no RecSource associated with ADC. So set the
       'DataMode property of ADC control as xdatUnbound
  ADC.DataMode = xdatUnbound
End Sub
Private Sub GrdOp_SetRecSource()
  '*** Set the Record Source of the Grid GrdOp
      'Default sort option is OpTc, OpYy, OpChr, OpNo, OpSr

  Dim ws_MainCnd As String, ws_SubCnd As String, ws_OrdBy As String
  
  ADC("wScOpTcFr").CmpStr = "OpTc >= ": ADC("wScOpTcTo").CmpStr = "OpTc <= "
  ADC("wScOpYyFr").CmpStr = "OpYy >= ": ADC("wScOpYyTo").CmpStr = "OpYy <= "
  ADC("wScOpChrFr").CmpStr = "OpChr >= ": ADC("wScOpChrTo").CmpStr = "OpChr <= "
  ADC("wScOpNoFr").CmpStr = "OpNo >= ": ADC("wScOpNoTo").CmpStr = "OpNo <= "
  ADC("wScOpSrFr").CmpStr = "OpSr >= ": ADC("wScOpSrTo").CmpStr = "OpSr <= "
  ADC("wScOmCmCdFr").CmpStr = "": ADC("wScOmCmCdTo").CmpStr = ""
  ADC("wScOmDtFr").CmpStr = "": ADC("wScOmDtTo").CmpStr = ""
  ADC("wScOdDelDtFr").CmpStr = "": ADC("wScOdDelDtTo").CmpStr = ""
  ADC("wScOdPrtCdFr").CmpStr = "": ADC("wScOdPrtCdTo").CmpStr = ""
  ADC("wScOpLineFr").CmpStr = "OpLine >= ": ADC("wScOpLineTo").CmpStr = "OpLine <= "    '*** (Jen 2.14)
  
  ws_MainCnd = ADC.RepCond
  ws_MainCnd = IIF(ws_MainCnd <> "", " and " + ws_MainCnd, "")
  
  ADC("wScOpTcFr").CmpStr = "": ADC("wScOpTcTo").CmpStr = ""
  ADC("wScOpYyFr").CmpStr = "": ADC("wScOpYyTo").CmpStr = ""
  ADC("wScOpChrFr").CmpStr = "": ADC("wScOpChrTo").CmpStr = ""
  ADC("wScOpNoFr").CmpStr = "": ADC("wScOpNoTo").CmpStr = ""
  ADC("wScOpSrFr").CmpStr = "": ADC("wScOpSrTo").CmpStr = ""
  ADC("wScOmCmCdFr").CmpStr = "OmCmCd >= ": ADC("wScOmCmCdTo").CmpStr = "OmCmCd <= "
  ADC("wScOmDtFr").CmpStr = "OmDt >= ": ADC("wScOmDtTo").CmpStr = "OmDt <= "
  ADC("wScOdDelDtFr").CmpStr = "OdDelDt >= ": ADC("wScOdDelDtTo").CmpStr = "OdDelDt <= "
  ADC("wScOdPrtCdFr").CmpStr = "OdPrtCd >= ": ADC("wScOdPrtCdTo").CmpStr = "OdPrtCd <= "
  ADC("wScOpLineFr").CmpStr = "": ADC("wScOpLineTo").CmpStr = ""      '*** (Jen 2.14)
  
  ws_SubCnd = ADC.RepCond
  ws_SubCnd = IIF(ws_SubCnd <> "", " and " + ws_SubCnd, "")
  ws_SubCnd = IIF(ws_SubCnd = "", "", " and exists(select 'x' from OrdDsg join OrdMst " + _
               "on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
               "where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr " + ws_SubCnd + " ) ")
               
  If ADC("wScSrt") = "" Or ADC("wScSrt") = "ORDSR" Then
    ms_OrdByFlds = "OdTc, OdYy, OdChr, OdNo, OdSr "
    ws_OrdBy = " Order By OpTc, OpYy, OpChr, OpNo, OpSr "
  ElseIf ADC("wScSrt") = "PRTCD" Then
    ms_OrdByFlds = "max(OdPrtCd), max(convert(varchar(8), OdDelDt, 112)), OdTc, OdYy, OdChr, OdNo, OdSr "
    ws_OrdBy = " Order By (Select (OdPrtCd+ convert(varchar(8), OdDelDt, 112)+ OdTc+ OdYy+ OdChr+ str(OdNo,10)+ str(OdSr,10)) " + _
               "from OrdDsg Where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr) "
  ElseIf ADC("wScSrt") = "DELDT" Then
    ms_OrdByFlds = "max(convert(varchar(8), OdDelDt, 112)), max(OdPrtCd), OdTc, OdYy, OdChr, OdNo, OdSr "
    ws_OrdBy = " Order By (Select (convert(varchar(8), OdDelDt, 112)+ OdPrtCd+ OdTc+ OdYy+ OdChr+ str(OdNo,10)+ str(OdSr,10)) " + _
               "from OrdDsg Where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr) "
  ElseIf ADC("wScSrt") = "DSGCD" Then
    ms_OrdByFlds = "max(OdDmCd), max(OdPrtCd), max(convert(varchar(8), OdDelDt, 112)), OdTc, OdYy, OdChr, OdNo, OdSr "
    ws_OrdBy = " Order By (Select (OdDmCd+ OdPrtCd+ convert(varchar(8), OdDelDt, 112)+ OdTc+ OdYy+ OdChr+ str(OdNo,10)+ str(OdSr,10)) " + _
               "from OrdDsg Where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr) "
  ElseIf ADC("wScSrt") = "SETTYP" Then
    ms_OrdByFlds = "max(OdWH), max(OdPrtCd), max(convert(varchar(8), OdDelDt, 112)), OdTc, OdYy, OdChr, OdNo, OdSr "
    ws_OrdBy = " Order By (Select (OdWH+ OdPrtCd+ convert(varchar(8), OdDelDt, 112)+ OdTc+ OdYy+ OdChr+ str(OdNo,10)+ str(OdSr,10)) " + _
               "from OrdDsg Where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr) "
  ElseIf ADC("wScSrt") = "LINE" Then
    ms_OrdByFlds = "max(OdLine), max(OdPrtCd), max(convert(varchar(8), OdDelDt, 112)), OdTc, OdYy, OdChr, OdNo, OdSr"
    ws_OrdBy = " Order By OpLine, (Select (OdPrtCd+ convert(varchar(8), OdDelDt, 112)+ OdTc+ OdYy+ OdChr+ str(OdNo,10)+ str(OdSr,10)) " + _
               "from OrdDsg Where OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
               "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr) "
  End If
  
  'ms_OrdByFlds = moCn.GetFldVal("Select PDesc225 from Param " + _
  '               "where PTyp= 'SRTOP' and PMCd= '" + ADC("wScSrt") + "'")
  'ws_OrdBy = " Order By " + ms_OrdByFlds
  
  GrdOp.RecSource = " Select * from OrdPln where OpCoCd= '" + gs_CoCd + "' " + ws_MainCnd + ws_SubCnd + ws_OrdBy
End Sub
Private Sub GrdOPD_SetRecSource()
  'If GrdOp.Row > 0 Then
    GrdOpd.RecSource = "Select * from OrdPlnDt " + _
                       "Where OpdCoCd= '" + gs_CoCd + "' " + _
                       "And OpdTc= '" + GrdOp.Value(GrdOp.Row, "OpTc") + "'" + _
                       "And OpdYy= '" + GrdOp.Value(GrdOp.Row, "OpYy") + "'" + _
                       "And OpdChr= '" + GrdOp.Value(GrdOp.Row, "OpChr") + "' " + _
                       "And OpdNo= " + CStr(GrdOp.Value(GrdOp.Row, "OpNo")) + _
                       "And OpdSr= " + CStr(GrdOp.Value(GrdOp.Row, "OpSr")) + " " + _
                       "Order By OpdPlnDt"
  'End If
End Sub
Private Sub CmdSc_Click()
  Call DispFra(Scop)
End Sub
Private Sub CmdResPln_Click()
  Call DispFra(ResPln)
  If FraResPln.Visible = False Then
    ADC("wRpOpTcFr") = "": ADC("wRpOpTcTo") = ""
    ADC("wRpOpYyFr") = "": ADC("wRpOpYyTo") = ""
    ADC("wRpOpChrFr") = "": ADC("wRpOpChrTo") = ""
    ADC("wRpOpNoFr") = "": ADC("wRpOpNoTo") = ""
    ADC("wRpOpSrFr") = "": ADC("wRpOpSrTo") = ""
    ADC("wRpOmCmCdFr") = "": ADC("wRpOmCmCdTo") = ""
    ADC("wRpOmDtFr") = "": ADC("wRpOmDtTo") = ""
    ADC("wRpOdDelDtFr") = "": ADC("wRpOdDelDtTo") = ""
    ADC("wRpOdPrtCdFr") = "": ADC("wRpOdPrtCdTo") = ""
    ADC("wRpOpLineFr") = "": ADC("wRpOpLineTo") = ""    '*** (Jen 2.14)
  Else
    ADC("wRpOpTcFr") = ms_wScOpTcFr: ADC("wRpOpTcTo") = ms_wScOpTcTo
    ADC("wRpOpYyFr") = ms_wScOpYyFr: ADC("wRpOpYyTo") = ms_wScOpYyTo
    ADC("wRpOpChrFr") = ms_wScOpChrFr: ADC("wRpOpChrTo") = ms_wScOpChrTo
    ADC("wRpOpNoFr") = mf_wScOpNoFr: ADC("wRpOpNoTo") = mf_wScOpNoTo
    ADC("wRpOpSrFr") = mf_wScOpSrFr: ADC("wRpOpSrTo") = mf_wScOpSrTo
    ADC("wRpOmCmCdFr") = ms_wScOmCmCdFr: ADC("wRpOmCmCdTo") = ms_wScOmCmCdTo
    ADC("wRpOmDtFr") = ms_wScOmDtFr: ADC("wRpOmDtTo") = ms_wScOmDtTo
    ADC("wRpOdDelDtFr") = ms_wScOdDelDtFr: ADC("wRpOdDelDtTo") = ms_wScOdDelDtTo
    ADC("wRpOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wRpOdPrtCdTo") = ms_wScOdPrtCdTo
    ADC("wRpOpLineFr") = ms_wScOpLineFr: ADC("wRpOpLineTo") = ms_wScOpLineTo    '*** (Jen 2.14)
  End If
End Sub
Private Sub CmdResPlnGo_Click()
  '*** (Bef 2.14)
  'code shifted to ResetPlan procedure
  '*** (Bef 2.14)
  
  Call ResetPlan    '*** (Jen 2.14)
  Call CmdResPln_Click
End Sub
Private Sub CmdUpdStat_Click()
  Call DispFra(UpdStat)
  If FraUpdStat.Visible = False Then
    ADC("wUsOpTcFr") = "": ADC("wUsOpTcTo") = ""
    ADC("wUsOpYyFr") = "": ADC("wUsOpYyTo") = ""
    ADC("wUsOpChrFr") = "": ADC("wUsOpChrTo") = ""
    ADC("wUsOpNoFr") = "": ADC("wUsOpNoTo") = ""
    ADC("wUsOpSrFr") = "": ADC("wUsOpSrTo") = ""
    ADC("wUsOmCmCdFr") = "": ADC("wUsOmCmCdTo") = ""
    ADC("wUsOmDtFr") = "": ADC("wUsOmDtTo") = ""
    ADC("wUsOdDelDtFr") = "": ADC("wUsOdDelDtTo") = ""
    ADC("wUsOdPrtCdFr") = "": ADC("wUsOdPrtCdTo") = ""
    ADC("wUsOpLineFr") = "": ADC("wUsOpLineTo") = ""    '*** (Jen 2.14)
  Else
    ADC("wUsOpTcFr") = ms_wScOpTcFr: ADC("wUsOpTcTo") = ms_wScOpTcTo
    ADC("wUsOpYyFr") = ms_wScOpYyFr: ADC("wUsOpYyTo") = ms_wScOpYyTo
    ADC("wUsOpChrFr") = ms_wScOpChrFr: ADC("wUsOpChrTo") = ms_wScOpChrTo
    ADC("wUsOpNoFr") = mf_wScOpNoFr: ADC("wUsOpNoTo") = mf_wScOpNoTo
    ADC("wUsOpSrFr") = mf_wScOpSrFr: ADC("wUsOpSrTo") = mf_wScOpSrTo
    ADC("wUsOmCmCdFr") = ms_wScOmCmCdFr: ADC("wUsOmCmCdTo") = ms_wScOmCmCdTo
    ADC("wUsOmDtFr") = ms_wScOmDtFr: ADC("wUsOmDtTo") = ms_wScOmDtTo
    ADC("wUsOdDelDtFr") = ms_wScOdDelDtFr: ADC("wUsOdDelDtTo") = ms_wScOdDelDtTo
    ADC("wUsOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wUsOdPrtCdTo") = ms_wScOdPrtCdTo
    
    '*** (Jen 2.14)
    ADC("wUsOpLineFr") = ms_wScOpLineFr: ADC("wUsOpLineTo") = ms_wScOpLineTo
    If ADC("wUsGetNewOrd") = "" Then ADC("wUsGetNewOrd") = "Y"
    '*** (Jen 2.14)
  End If
End Sub
Private Sub CmdUpdStatGo_Click()
  '*** Updates the Ord Plan status
  Call UpdateStatus(False)  '*** (Jen 2.14)
  'Call UpdateStatus        '*** (Bef 2.14)
  Call CmdUpdStat_Click
End Sub
Private Sub CmdGenPln_Click()
  Call DispFra(GenPln)
  If FraGenPln.Visible = False Then
    ADC("wGpOpTcFr") = "": ADC("wGpOpTcTo") = ""
    ADC("wGpOpYyFr") = "": ADC("wGpOpYyTo") = ""
    ADC("wGpOpChrFr") = "": ADC("wGpOpChrTo") = ""
    ADC("wGpOpNoFr") = "": ADC("wGpOpNoTo") = ""
    ADC("wGpOpSrFr") = "": ADC("wGpOpSrTo") = ""
    ADC("wGpOmCmCdFr") = "": ADC("wGpOmCmCdTo") = ""
    ADC("wGpOmDtFr") = "": ADC("wGpOmDtTo") = ""
    ADC("wGpOdDelDtFr") = "": ADC("wGpOdDelDtTo") = ""
    ADC("wGpOdPrtCdFr") = "": ADC("wGpOdPrtCdTo") = ""
    ADC("wGpOpLineFr") = "": ADC("wGpOpLineTo") = ""    '*** (Jen 2.14)
  Else
    ADC("wGpOpTcFr") = ms_wScOpTcFr: ADC("wGpOpTcTo") = ms_wScOpTcTo
    ADC("wGpOpYyFr") = ms_wScOpYyFr: ADC("wGpOpYyTo") = ms_wScOpYyTo
    ADC("wGpOpChrFr") = ms_wScOpChrFr: ADC("wGpOpChrTo") = ms_wScOpChrTo
    ADC("wGpOpNoFr") = mf_wScOpNoFr: ADC("wGpOpNoTo") = mf_wScOpNoTo
    ADC("wGpOpSrFr") = mf_wScOpSrFr: ADC("wGpOpSrTo") = mf_wScOpSrTo
    ADC("wGpOmCmCdFr") = ms_wScOmCmCdFr: ADC("wGpOmCmCdTo") = ms_wScOmCmCdTo
    ADC("wGpOmDtFr") = ms_wScOmDtFr: ADC("wGpOmDtTo") = ms_wScOmDtTo
    ADC("wGpOdDelDtFr") = ms_wScOdDelDtFr: ADC("wGpOdDelDtTo") = ms_wScOdDelDtTo
    ADC("wGpOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wGpOdPrtCdTo") = ms_wScOdPrtCdTo
    ADC("wGpOpLineFr") = ms_wScOpLineFr: ADC("wGpOpLineTo") = ms_wScOpLineTo    '*** (Jen 2.14)
    
    ADC("wGpGapDysP1") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P1' and PSCd= ''")
    ADC("wGpGapDysP2") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P2' and PSCd= ''")
    ADC("wGpGapDysP3") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P3' and PSCd= ''")
    ADC("wGpGapDysP4") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P4' and PSCd= ''")
    ADC("wGpGapDysP5") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P5' and PSCd= ''")
  End If
End Sub
Private Sub CmdGenPlnGo_Click()
  '*** Call procedure to generate the plan
  Call GeneratePlan(False)  '*** (Jen 2.14)
  'Call GeneratePlan        '*** (Bef 2.14)
  Call CmdGenPln_Click
End Sub
Private Sub CmdAutoGenPln_Click()
  '*** (Jen 2.14)
  Call DispFra(AutoGenPln)
  If FraAutoGenPln.Visible = False Then
    ADC("wAgOpTcFr") = "": ADC("wAgOpTcTo") = ""
    ADC("wAgOpYyFr") = "": ADC("wAgOpYyTo") = ""
    ADC("wAgOpChrFr") = "": ADC("wAgOpChrTo") = ""
    ADC("wAgOpNoFr") = "": ADC("wAgOpNoTo") = ""
    ADC("wAgOpSrFr") = "": ADC("wAgOpSrTo") = ""
    ADC("wAgOmCmCdFr") = "": ADC("wAgOmCmCdTo") = ""
    ADC("wAgOmDtFr") = "": ADC("wAgOmDtTo") = ""
    ADC("wAgOdDelDtFr") = "": ADC("wAgOdDelDtTo") = ""
    ADC("wAgOdPrtCdFr") = "": ADC("wAgOdPrtCdTo") = ""
    ADC("wAgOpLineFr") = "": ADC("wAgOpLineTo") = ""
  Else
    ADC("wAgOpTcFr") = ms_wScOpTcFr: ADC("wAgOpTcTo") = ms_wScOpTcTo
    ADC("wAgOpYyFr") = ms_wScOpYyFr: ADC("wAgOpYyTo") = ms_wScOpYyTo
    ADC("wAgOpChrFr") = ms_wScOpChrFr: ADC("wAgOpChrTo") = ms_wScOpChrTo
    ADC("wAgOpNoFr") = mf_wScOpNoFr: ADC("wAgOpNoTo") = mf_wScOpNoTo
    ADC("wAgOpSrFr") = mf_wScOpSrFr: ADC("wAgOpSrTo") = mf_wScOpSrTo
    ADC("wAgOmCmCdFr") = ms_wScOmCmCdFr: ADC("wAgOmCmCdTo") = ms_wScOmCmCdTo
    ADC("wAgOmDtFr") = ms_wScOmDtFr: ADC("wAgOmDtTo") = ms_wScOmDtTo
    ADC("wAgOdDelDtFr") = ms_wScOdDelDtFr: ADC("wAgOdDelDtTo") = ms_wScOdDelDtTo
    ADC("wAgOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wAgOdPrtCdTo") = ms_wScOdPrtCdTo
    ADC("wAgOpLineFr") = ms_wScOpLineFr: ADC("wAgOpLineTo") = ms_wScOpLineTo
    
    ADC("wAgGapDysP1") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P1' and PSCd= ''")
    ADC("wAgGapDysP2") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P2' and PSCd= ''")
    ADC("wAgGapDysP3") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P3' and PSCd= ''")
    ADC("wAgGapDysP4") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P4' and PSCd= ''")
    ADC("wAgGapDysP5") = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and PMCd= 'P5' and PSCd= ''")
    
    If ADC("wAgGetNewOrd") = "" Then ADC("wAgGetNewOrd") = "Y"
  End If
  '*** (Jen 2.14)
End Sub
Private Sub CmdAutoGenPlnGo_Click()
  '*** (Jen 2.14)
  '*** Call procedure to Auto-generate Plan For The Specified Scope After The Specified Interval
  
  Dim wi_PlnDays As Integer
  If ADC("wAgGetNewOrd") = "" Then DispMsg "Specify Whether To Introduce New Orders Belonging To The Given Scope For Planning", etError: Exit Sub
  
  wi_PlnDays = moCn.GetFldVal("Select PNum From Param where PTyp= 'PLNDAYS' and PMCd= 'PLNDAYS'")
  If wi_PlnDays <= 0 Then DispMsg "Number Of Days For Which Plan Has To Be Generated Should Be Specified In Param('PLNDAYS')", etError: Exit Sub
  
  If ADC("wAgIntr") = 0 Then
    TimerAutoGen.Enabled = False
    TimerAutoGen.Interval = 0
    DispMsg "Plan Auto-Generation Has Been Cancelled", etError
    Call CmdAutoGenPln_Click
    Exit Sub
  Else
    ms_wScOpTcFr = ADC("wAgOpTcFr"): ms_wScOpTcTo = ADC("wAgOpTcTo")
    ms_wScOpYyFr = ADC("wAgOpYyFr"): ms_wScOpYyTo = ADC("wAgOpYyTo")
    ms_wScOpChrFr = ADC("wAgOpChrFr"): ms_wScOpChrTo = ADC("wAgOpChrTo")
    mf_wScOpNoFr = ADC("wAgOpNoFr"): mf_wScOpNoTo = ADC("wAgOpNoTo")
    mf_wScOpSrFr = ADC("wAgOpSrFr"): mf_wScOpSrTo = ADC("wAgOpSrTo")
    ms_wScOmCmCdFr = ADC("wAgOmCmCdFr"): ms_wScOmCmCdTo = ADC("wAgOmCmCdTo")
    ms_wScOmDtFr = ADC("wAgOmDtFr"): ms_wScOmDtTo = ADC("wAgOmDtTo")
    ms_wScOdDelDtFr = ADC("wAgOdDelDtFr"): ms_wScOdDelDtTo = ADC("wAgOdDelDtTo")
    ms_wScOdPrtCdFr = ADC("wAgOdPrtCdFr"): ms_wScOdPrtCdTo = ADC("wAgOdPrtCdTo")
    ms_wScOpLineFr = ADC("wAgOpLineFr"): ms_wScOpLineTo = ADC("wAgOpLineTo")
  
    '*** (Bef 3.1.2)
    'TimerAutoGen.Interval = ADC("wAgIntr") * 1000 * 60      '*** 1 second= 1000 milliseconds
    '*** (Bef 3.1.2)
    
    '*** (Jen 3.1.2)
    mf_Time = ADC("wAgIntr") * 1000 * 60     '*** 1 second= 1000 milliseconds
    TimerAutoGen.Interval = IIF(mf_Time < 50000, mf_Time, 50000)
    mf_Time = mf_Time - TimerAutoGen.Interval
    '*** (Jen 3.1.2)
    
    TimerAutoGen.Enabled = True
  End If
  '*** (Jen 2.14)
End Sub
Private Sub TimerAutoGen_Timer()
  '*** (Jen 3.1.2)
  If mf_Time > 0 Then
    TimerAutoGen.Interval = IIF(mf_Time < 50000, mf_Time, 50000)
    mf_Time = mf_Time - TimerAutoGen.Interval
    Exit Sub
  'Else
  '  MsgBox "Set Time"
  End If
  '*** (Jen 3.1.2)
  
  '*** (Jen 2.14)
  ADC("wAgOpTcFr") = ms_wScOpTcFr: ADC("wAgOpTcTo") = ms_wScOpTcTo
  ADC("wAgOpYyFr") = ms_wScOpYyFr: ADC("wAgOpYyTo") = ms_wScOpYyTo
  ADC("wAgOpChrFr") = ms_wScOpChrFr: ADC("wAgOpChrTo") = ms_wScOpChrTo
  ADC("wAgOpNoFr") = mf_wScOpNoFr: ADC("wAgOpNoTo") = mf_wScOpNoTo
  ADC("wAgOpSrFr") = mf_wScOpSrFr: ADC("wAgOpSrTo") = mf_wScOpSrTo
  ADC("wAgOmCmCdFr") = ms_wScOmCmCdFr: ADC("wAgOmCmCdTo") = ms_wScOmCmCdTo
  ADC("wAgOmDtFr") = ms_wScOmDtFr: ADC("wAgOmDtTo") = ms_wScOmDtTo
  ADC("wAgOdDelDtFr") = ms_wScOdDelDtFr: ADC("wAgOdDelDtTo") = ms_wScOdDelDtTo
  ADC("wAgOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wAgOdPrtCdTo") = ms_wScOdPrtCdTo
  ADC("wAgOpLineFr") = ms_wScOpLineFr: ADC("wAgOpLineTo") = ms_wScOpLineTo
  
  Call ResetPlan
  
  Call UpdateStatus(True)
  
  ADC("wAgOpTcFr") = ms_wScOpTcFr: ADC("wAgOpTcTo") = ms_wScOpTcTo
  ADC("wAgOpYyFr") = ms_wScOpYyFr: ADC("wAgOpYyTo") = ms_wScOpYyTo
  ADC("wAgOpChrFr") = ms_wScOpChrFr: ADC("wAgOpChrTo") = ms_wScOpChrTo
  ADC("wAgOpNoFr") = mf_wScOpNoFr: ADC("wAgOpNoTo") = mf_wScOpNoTo
  ADC("wAgOpSrFr") = mf_wScOpSrFr: ADC("wAgOpSrTo") = mf_wScOpSrTo
  ADC("wAgOmCmCdFr") = ms_wScOmCmCdFr: ADC("wAgOmCmCdTo") = ms_wScOmCmCdTo
  ADC("wAgOmDtFr") = ms_wScOmDtFr: ADC("wAgOmDtTo") = ms_wScOmDtTo
  ADC("wAgOdDelDtFr") = ms_wScOdDelDtFr: ADC("wAgOdDelDtTo") = ms_wScOdDelDtTo
  ADC("wAgOdPrtCdFr") = ms_wScOdPrtCdFr: ADC("wAgOdPrtCdTo") = ms_wScOdPrtCdTo
  ADC("wAgOpLineFr") = ms_wScOpLineFr: ADC("wAgOpLineTo") = ms_wScOpLineTo
  
  Call GeneratePlan(True)
  TimerAutoGen.Enabled = False
  TimerAutoGen.Interval = 0
  If FraAutoGenPln.Visible = True Then
    Call CmdAutoGenPln_Click
  End If
  
  ADC("wAgOpTcFr") = "": ADC("wAgOpTcTo") = ""
  ADC("wAgOpYyFr") = "": ADC("wAgOpYyTo") = ""
  ADC("wAgOpChrFr") = "": ADC("wAgOpChrTo") = ""
  ADC("wAgOpNoFr") = "": ADC("wAgOpNoTo") = ""
  ADC("wAgOpSrFr") = "": ADC("wAgOpSrTo") = ""
  ADC("wAgOmCmCdFr") = "": ADC("wAgOmCmCdTo") = ""
  ADC("wAgOmDtFr") = "": ADC("wAgOmDtTo") = ""
  ADC("wAgOdDelDtFr") = "": ADC("wAgOdDelDtTo") = ""
  ADC("wAgOdPrtCdFr") = "": ADC("wAgOdPrtCdTo") = ""
  ADC("wAgOpLineFr") = "": ADC("wAgOpLineTo") = ""
  '*** (Jen 2.14)
End Sub
Private Sub CmdShwLd_Click()
  Call DispFra(ShwLd)
  If FraShwLd.Visible = True Then
    ADC("wSlPlnDtFr") = moCn.SrvrDate
    ADC("wSlPlnDtTo") = DateAdd("d", 6, moCn.SrvrDate)
    If GrdOp.Row > 0 Then ADC("wSlLine") = GrdOp.Value(GrdOp.Row, "OpLine")
    If ADC("wSlLine") <> "" Then Call ShowLoad
  End If
End Sub
Private Sub CmdShwLdGo_Click()
  '*** call procedure to sow the Ld/ Cap
  Call ShowLoad
End Sub
Private Sub CmdDsgDet_Click()
  Call DispFra(DsgDet)
  If FraDsgDet.Visible = True Then Call DesignDetails
End Sub
Private Sub CmdRedoPln_Click()
  If GrdOpd.Rows - 1 = 0 Then Exit Sub
  If FraRedoPln.Visible = False Then
    If GrdOpd.HalfValidate Then GrdOpd.SetFocus: Exit Sub
  End If
  Call DispFra(RedoPln)
End Sub
Private Sub CmdRedoPlnGo_Click()
  Call PlanRedo
  Call CmdRedoPln_Click
End Sub
Private Sub ResetPlan()
  '*** (Jen 2.14)
  Dim ws_Cnd As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  moCn.Execute ("Delete OrdPlnDt from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc " + _
                "and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo where OpdCoCd= OdCoCd and " + _
                "OpdTc= OdTc and OpdYy= OdYy and OpdChr= OdChr and OpdNo= OdNo and " + _
                "OpdSr= OdSr and OdCoCd= '" + gs_CoCd + "' " + ws_Cnd)
  
  If GrdOpd.Rows - 1 > 0 Then
    GrdOpd.Init
    GrdOpd.RefreshData
  End If
  '*** (Jen 2.14)
End Sub
Private Sub UpdateStatus(ByVal pb_AutoGen As Boolean)
  '*** ByVal pb_AutoGen As Boolean '*** (added Jen 2.14)
  
  Dim ws_Cnd As String, wRs_OrdPln As MDORowSet, wi_i As Integer
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  
  If moCn.RecSeek("Select 'x' from OrdPlnDt join OrdDsg on OdCoCd= OpdCoCd and OdTC= OpdTc and " + _
                  "OdYy= OpdYy and OdChr= OpdChr and OdNo= OpdNo and OdSr= OpdSr " + _
                  "join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
                  "OmChr= OdChr and OmNo= OdNo where OpdCoCd= '" + gs_CoCd + "' " + ws_Cnd) Then
    DispMsg "Select The 'Reset Plan' Option Before 'Update Status' As Rec Exists In OrdPlanDate", etError: Exit Sub
  End If
  
  '*** (Jen 2.14)
  If pb_AutoGen = True Then
    If ADC("wAgGetNewOrd") = "" Then DispMsg "Specify Whether To Introduce New Orders Belonging To The Given Scope For Planning", etError: Exit Sub
  Else
    If ADC("wUsGetNewOrd") = "" Then DispMsg "Specify Whether To Introduce New Orders Belonging To The Given Scope For Planning", etError: Exit Sub
  End If
  '*** (Jen 2.14)
  
  ''If moCn.RecSeek("Select 'x' from OrdPlnDt where OpdCoCd= '" + gs_CoCd + "' and " + _
                  "exists(Select 'x' from OrdDsg join OrdMst on " + _
                  "OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
                  "where OdCoCd= OpdCoCd and OdTC= OpdTc and OdYy= OpdYy and OdChr= OpdChr " + _
                  "and OdNo= OpdNo and OdSr= OpdSr and OdPrdQty> 0 " + _
                  "and Str(OdPrdQty, 16, 1)= Str(OdFgQty, 16, 1) " + ws_Cnd + ") ") Then
  ''  DispMsg "Select The 'Reset Plan' Option Before 'Update Status' As Rec Exists In OrdPlanDate", etError: Exit Sub
  ''End If
  
  moCn.Execute ("Delete from OrdPln where OpCoCd= '" + gs_CoCd + "' and " + _
                "exists(Select 'x' from OrdDsg join OrdMst on " + _
                "OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
                "where OdCoCd= OpCoCd and OdTC= OpTc and OdYy= OpYy and OdChr= OpChr " + _
                "and OdNo= OpNo and OdSr= OpSr and OdPrdQty> 0 " + _
                "and Str(OdPrdQty, 16, 1)= Str(OdFgQty, 16, 1) " + ws_Cnd + ") ")
  
  '*** This statement cannot be run as aggregate functions cannot be used in the set list of an update statement
  '''''moCn.Execute ("Update OrdPln set OpLine= max(OdLine), OpWDiaAvlblDt= max(OdWDiaAvlblDt), " + _
                "OpHDiaAvlblDt= max(OdHDiaAvlblDt), " + _
                "OpWipwSet= sum(case when IsNull(BPrdStg, '')= 'WSET' then IsNull(BQty, 0) else 0 end), " + _
                "OpWipCas= sum(case when IsNull(BPrdStg, '')= 'CAS' then IsNull(BQty, 0) else 0 end), " + _
                "OpWipFin= sum(case when IsNull(BPrdStg, '')= 'FIN' then IsNull(BQty, 0) else 0 end), " + _
                "OpWiphSet= sum(case when IsNull(BPrdStg, '')= 'HSET' then IsNull(BQty, 0) else 0 end), " + _
                "OpWipEi= sum(case when IsNull(BPrdStg, '')= 'EI' then IsNull(BQty, 0) else 0 end) " + _
                "from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
                "OmChr= OdChr and OmNo= OdNo left outer join Bag on BCoCd= OpCoCd and " + _
                "BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and BOdNo= OpNo and BOdSr= OpSr " + _
                "and BCls= 'N' and BLoc<> 'PFG' " + _
                "where OpCoCd= OdCoCd and OpTc= OdTc and OpYy= OdYy and " + _
                "OpChr= OdChr and OpNo= OdNo and OpSr= OdSr and OdCoCd= '" + gs_CoCd + "' " + ws_Cnd + _
                " Group By OdTc, OdYy, OdChr, OdNo, OdSr ")
  '*** This statement cannot be run as aggregate functions cannot be used in the set list of an update statement
    
      
  '*** (Bef 2.14) (This code works too but the Bag table is accessed many times)
  ''****** (Jen 2.12 patch 2) Flute bags are not considered ******/
  'moCn.Execute ("Update OrdPln set /* OpLine= OdLine, OpWDiaAvlblDt= OdWDiaAvlblDt, " + _
  '              "OpHDiaAvlblDt= OdHDiaAvlblDt,*/ " + _
  '              "OpWipwSet= IsNull((Select sum(BQty) from Bag " + _
  '              "   where BCoCd= OpCoCd and BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and " + _
  '              "   BOdNo= OpNo and BOdSr= OpSr and BPrdStg= 'WSET' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N') ), 0), " + _
  '              "OpWipCas= IsNull((Select sum(BQty) from Bag " + _
  '              "   where BCoCd= OpCoCd and BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and " + _
  '              "   BOdNo= OpNo and BOdSr= OpSr and BPrdStg= 'CAS' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N') ), 0), " + _
  '              "OpWipFin= IsNull((Select sum(BQty) from Bag " + _
  '              "   where BCoCd= OpCoCd and BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and " + _
  '              "   BOdNo= OpNo and BOdSr= OpSr and BPrdStg= 'FIN' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N') ), 0), " + _
  '              "OpWiphSet= IsNull((Select sum(BQty) from Bag " + _
  '              "   where BCoCd= OpCoCd and BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and " + _
  '              "   BOdNo= OpNo and BOdSr= OpSr and BPrdStg= 'HSET' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N') ), 0), " + _
  '              "OpWipEi= IsNull((Select sum(BQty) from Bag " + _
  '              "   where BCoCd= OpCoCd and BOdTc= OpTc and BOdYy= OpYy and BOdChr= OpChr and " + _
  '              "   BOdNo= OpNo and BOdSr= OpSr and BPrdStg= 'EI' and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N') ), 0) " + _
  '              "from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
  '              "OmChr= OdChr and OmNo= OdNo where OpCoCd= OdCoCd and OpTc= OdTc and OpYy= OdYy and " + _
  '              "OpChr= OdChr and OpNo= OdNo and OpSr= OdSr and OdCoCd= '" + gs_CoCd + "' " + ws_Cnd)
  '*** (Bef 2.14)
                
                
  '*** (Jen 2.14)
  '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
  moCn.Execute ("Update OrdPln set /* OpLine= OdLine, OpWDiaAvlblDt= OdWDiaAvlblDt, " + _
                "OpHDiaAvlblDt= OdHDiaAvlblDt,*/ " + _
                "OpWipP1= qOpWipP1, OpWipP2= qOpWipP2, OpWipP3= qOpWipP3, OpWipP4= qOpWipP4, " + _
                "OpWipP5= qOpWipP5 " + _
                "From " + _
                "(Select OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, " + _
                "     sum(case when IsNull(BPrdStg, '')= 'P1' then IsNull(BQty, 0) else 0 end) as qOpWipP1, " + _
                "     sum(case when IsNull(BPrdStg, '')= 'P2' then IsNull(BQty, 0) else 0 end) as qOpWipP2, " + _
                "     sum(case when IsNull(BPrdStg, '')= 'P3' then IsNull(BQty, 0) else 0 end) as qOpWipP3, " + _
                "     sum(case when IsNull(BPrdStg, '')= 'P4' then IsNull(BQty, 0) else 0 end) as qOpWipP4, " + _
                "     sum(case when IsNull(BPrdStg, '')= 'P5' then IsNull(BQty, 0) else 0 end) as qOpWipP5 " + _
                "     from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
                "     left outer join Bag On BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr " + _
                "         and BOdNo= OdNo and BOdSr= OdSr " + _
                "         and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') " + _
                "     where OdCoCd= '" + gs_CoCd + "' " + ws_Cnd + _
                "     group By OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr) Tmp " + _
                "where OpCoCd= OdCoCd and OpTc= OdTc and OpYy= OdYy and " + _
                "OpChr= OdChr and OpNo= OdNo and OpSr= OdSr and OpCoCd= '" + gs_CoCd + "' ")
    
               '*** (Jen 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
               '*** (Jen 2.14 Next)
               
               '*** (Bef 2.14 Next)
               'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
               '*** (Bef 2.14 Next)
    
  If pb_AutoGen = True Then
    ADC("wScOpTcFr") = ADC("wAgOpTcFr"): ADC("wScOpTcTo") = ADC("wAgOpTcTo")
    ADC("wScOpYyFr") = ADC("wAgOpYyFr"): ADC("wScOpYyTo") = ADC("wAgOpYyTo")
    ADC("wScOpChrFr") = ADC("wAgOpChrFr"): ADC("wScOpChrTo") = ADC("wAgOpChrTo")
    ADC("wScOpNoFr") = ADC("wAgOpNoFr"): ADC("wScOpNoTo") = ADC("wAgOpNoTo")
    ADC("wScOpSrFr") = ADC("wAgOpSrFr"): ADC("wScOpSrTo") = ADC("wAgOpSrTo")
    ADC("wScOmCmCdFr") = ADC("wAgOmCmCdFr"): ADC("wScOmCmCdTo") = ADC("wAgOmCmCdTo")
    ADC("wScOmDtFr") = ADC("wAgOmDtFr"): ADC("wScOmDtTo") = ADC("wAgOmDtTo")
    ADC("wScOdDelDtFr") = ADC("wAgOdDelDtFr"): ADC("wScOdDelDtTo") = ADC("wAgOdDelDtTo")
    ADC("wScOdPrtCdFr") = ADC("wAgOdPrtCdFr"): ADC("wScOdPrtCdTo") = ADC("wAgOdPrtCdTo")
    ADC("wScOpLineFr") = ADC("wAgOpLineFr"): ADC("wScOpLineTo") = ADC("wAgOpLineTo")
  
    ADC("wAgOpTcFr") = "": ADC("wAgOpTcTo") = ""
    ADC("wAgOpYyFr") = "": ADC("wAgOpYyTo") = ""
    ADC("wAgOpChrFr") = "": ADC("wAgOpChrTo") = ""
    ADC("wAgOpNoFr") = "": ADC("wAgOpNoTo") = ""
    ADC("wAgOpSrFr") = "": ADC("wAgOpSrTo") = ""
    ADC("wAgOmCmCdFr") = "": ADC("wAgOmCmCdTo") = ""
    ADC("wAgOmDtFr") = "": ADC("wAgOmDtTo") = ""
    ADC("wAgOdDelDtFr") = "": ADC("wAgOdDelDtTo") = ""
    ADC("wAgOdPrtCdFr") = "": ADC("wAgOdPrtCdTo") = ""
    ADC("wAgOpLineFr") = "": ADC("wAgOpLineTo") = ""
  Else
  '*** (Jen 2.14)
  
    ADC("wScOpTcFr") = ADC("wUsOpTcFr"): ADC("wScOpTcTo") = ADC("wUsOpTcTo")
    ADC("wScOpYyFr") = ADC("wUsOpYyFr"): ADC("wScOpYyTo") = ADC("wUsOpYyTo")
    ADC("wScOpChrFr") = ADC("wUsOpChrFr"): ADC("wScOpChrTo") = ADC("wUsOpChrTo")
    ADC("wScOpNoFr") = ADC("wUsOpNoFr"): ADC("wScOpNoTo") = ADC("wUsOpNoTo")
    ADC("wScOpSrFr") = ADC("wUsOpSrFr"): ADC("wScOpSrTo") = ADC("wUsOpSrTo")
    ADC("wScOmCmCdFr") = ADC("wUsOmCmCdFr"): ADC("wScOmCmCdTo") = ADC("wUsOmCmCdTo")
    ADC("wScOmDtFr") = ADC("wUsOmDtFr"): ADC("wScOmDtTo") = ADC("wUsOmDtTo")
    ADC("wScOdDelDtFr") = ADC("wUsOdDelDtFr"): ADC("wScOdDelDtTo") = ADC("wUsOdDelDtTo")
    ADC("wScOdPrtCdFr") = ADC("wUsOdPrtCdFr"): ADC("wScOdPrtCdTo") = ADC("wUsOdPrtCdTo")
    ADC("wScOpLineFr") = ADC("wUsOpLineFr"): ADC("wScOpLineTo") = ADC("wUsOpLineTo")    '*** (Jen 2.14)
  
    ADC("wUsOpTcFr") = "": ADC("wUsOpTcTo") = ""
    ADC("wUsOpYyFr") = "": ADC("wUsOpYyTo") = ""
    ADC("wUsOpChrFr") = "": ADC("wUsOpChrTo") = ""
    ADC("wUsOpNoFr") = "": ADC("wUsOpNoTo") = ""
    ADC("wUsOpSrFr") = "": ADC("wUsOpSrTo") = ""
    ADC("wUsOmCmCdFr") = "": ADC("wUsOmCmCdTo") = ""
    ADC("wUsOmDtFr") = "": ADC("wUsOmDtTo") = ""
    ADC("wUsOdDelDtFr") = "": ADC("wUsOdDelDtTo") = ""
    ADC("wUsOdPrtCdFr") = "": ADC("wUsOdPrtCdTo") = ""
    ADC("wUsOpLineFr") = "": ADC("wUsOpLineTo") = ""      '*** (Jen 2.14)
  
  '*** (Jen 2.14)
  End If
  '*** (Jen 2.14)
                
  GrdOp.Init
  GrdOp.RefreshData
  
  '*** Bef 16/01/04
  'ADC("wScOpTcFr") = "": ADC("wScOpTcTo") = ""
  'ADC("wScOpYyFr") = "": ADC("wScOpYyTo") = ""
  'ADC("wScOpChrFr") = "": ADC("wScOpChrTo") = ""
  'ADC("wScOpNoFr") = "": ADC("wScOpNoTo") = ""
  'ADC("wScOpSrFr") = "": ADC("wScOpSrTo") = ""
  'ADC("wScOmCmCdFr") = "": ADC("wScOmCmCdTo") = ""
  'ADC("wScOmDtFr") = "": ADC("wScOmDtTo") = ""
  'ADC("wScOdDelDtFr") = "": ADC("wScOdDelDtTo") = ""
  'ADC("wScOdPrtCdFr") = "": ADC("wScOdPrtCdTo") = ""
  '*** Bef 16/01/04
  
  '*** (Jen 2.14)
  Dim ws_Opt As String
  If pb_AutoGen = True Then
    ws_Opt = ADC("wAgGetNewOrd")
  Else
    ws_Opt = ADC("wUsGetNewOrd")
  End If
  
  If ws_Opt = "Y" Then
  '*** (Jen 2.14)
  
    '****** (Jen 2.12 patch 2) Flute bags are not considered ******/
    Set wRs_OrdPln = moCn.OpenRes("Select OdTc, OdYy, OdChr, OdNo, OdSr, max(OdDelDt) as OdDelDt, " + _
                     "max(OdLine) as OdLine, max(OdWDiaAvlblDt) as OdWDiaAvlblDt, " + _
                     "max(OdHDiaAvlblDt) as OdHDiaAvlblDt, max(OdDmCd) as OdDmCd, max(OdWH) as OdWH, " + _
                     "max(OdPrtCd) as OdPrtCd, max(OdPrdQty) as OdPrdQty, " + _
                     "max(OdPrdQty- OdFgQty) as qBalQty, sum(IsNull(BQty, 0)) as qFlrQty, " + _
                     "sum(case when IsNull(BPrdStg, '') = 'P1' then IsNull(BQty, 0) else 0 end) as qWipP1, " + _
                     "sum(case when IsNull(BPrdStg, '') = 'P2' then IsNull(BQty, 0) else 0 end) as qWipP2, " + _
                     "sum(case when IsNull(BPrdStg, '') = 'P3' then IsNull(BQty, 0) else 0 end) as qWipP3, " + _
                     "sum(case when IsNull(BPrdStg, '') = 'P4' then IsNull(BQty, 0) else 0 end) as qWipP4, " + _
                     "sum(case when IsNull(BPrdStg, '') = 'P5' then IsNull(BQty, 0) else 0 end) as qWipP5 " + _
                     "from OrdDsg join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
                     "OmChr= OdChr and OmNo= OdNo " + _
                     "left outer join Bag on BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy " + _
                     "and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr and BCls= 'N' and BLoc<> 'PFG' " + _
                     "and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P') " + _
                     "where OdCoCd= '" + gs_CoCd + "' and OdPrdQty> 0 and " + _
                     "(Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= 'SO' and " + _
                     "Str(OdPrdQty, 16, 1)<> Str(OdFgQty, 16, 1) and " + _
                     "not exists(Select 'x' from OrdPln where OpCoCd= OdCoCd and OpTc= OdTc " + _
                     "and OpYy= OdYy and OpChr= OdChr and OpNo= OdNo and OpSr= OdSr) " + ws_Cnd + " " + _
                     "Group By OdTc, OdYy, OdChr, OdNo, OdSr " + _
                     IIF(Trim(ms_OrdByFlds) <> "", " Order By " + ms_OrdByFlds, ""))
              
                     '*** Bef 16/01/04 Order By OdTc, OdYy, OdChr, OdNo, OdSr
                     
                    '*** (Jen 2.14 Next)
                    'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'P')
                    '*** (Jen 2.14 Next)
                    
                    '*** (Bef 2.14 Next)
                    'and exists(Select 'x' from Param where PTyp= 'BCHR' and PMCd= BChr and PValue= 'N')
                    '*** (Bef 2.14 Next)
    With GrdOp
      .StartCopy
      Do While Not (wRs_OrdPln.EOF Or wRs_OrdPln.BOF)
        .AddItem
        wi_i = .Rows - 1
        .Value(wi_i, "OpTc") = wRs_OrdPln!OdTc
        .Value(wi_i, "OpYy") = wRs_OrdPln!OdYy
        .Value(wi_i, "OpChr") = wRs_OrdPln!OdChr
        .Value(wi_i, "OpNo") = wRs_OrdPln!OdNo
        .Value(wi_i, "OpSr") = wRs_OrdPln!OdSr
        .Value(wi_i, "wOdDelDt") = wRs_OrdPln!OdDelDt
        .Value(wi_i, "OpLine") = wRs_OrdPln!OdLine
        .Value(wi_i, "OpWDiaAvlblDt") = wRs_OrdPln!OdWDiaAvlblDt
        .Value(wi_i, "OpHDiaAvlblDt") = wRs_OrdPln!OdHDiaAvlblDt
        .Value(wi_i, "wOdDmCd") = wRs_OrdPln!OdDmCd
        .Value(wi_i, "wOdWH") = wRs_OrdPln!OdWH
        .Value(wi_i, "wOdPrtCd") = wRs_OrdPln!OdPrtCd
        .Value(wi_i, "wOdPrdQty") = wRs_OrdPln!OdPrdQty
        .Value(wi_i, "wBalQty") = wRs_OrdPln!qBalQty
        .Value(wi_i, "wFlrQty") = wRs_OrdPln!qFlrQty
        .Value(wi_i, "OpWipP1") = wRs_OrdPln!qWipP1
        .Value(wi_i, "OpWipP2") = wRs_OrdPln!qWipP2
        .Value(wi_i, "OpWipP3") = wRs_OrdPln!qWipP3
        .Value(wi_i, "OpWipP4") = wRs_OrdPln!qWipP4
        .Value(wi_i, "OpWipP5") = wRs_OrdPln!qWipP5
        .SaveRec (wi_i)
        wRs_OrdPln.MoveNext
      Loop
      .EndCopy
    End With
    
  '*** (Jen 2.14)
  End If
  '*** (Jen 2.14)
  
  Set wRs_OrdPln = Nothing
  
  '*** 16/01/04
  GrdOp.Init
  GrdOp.RefreshData
  ADC("wScOpTcFr") = "": ADC("wScOpTcTo") = ""
  ADC("wScOpYyFr") = "": ADC("wScOpYyTo") = ""
  ADC("wScOpChrFr") = "": ADC("wScOpChrTo") = ""
  ADC("wScOpNoFr") = "": ADC("wScOpNoTo") = ""
  ADC("wScOpSrFr") = "": ADC("wScOpSrTo") = ""
  ADC("wScOmCmCdFr") = "": ADC("wScOmCmCdTo") = ""
  ADC("wScOmDtFr") = "": ADC("wScOmDtTo") = ""
  ADC("wScOdDelDtFr") = "": ADC("wScOdDelDtTo") = ""
  ADC("wScOdPrtCdFr") = "": ADC("wScOdPrtCdTo") = ""
  ADC("wScOpLineFr") = "": ADC("wScOpLineTo") = ""    '*** (Jen 2.14)
  '*** 16/01/04
  
End Sub
Private Sub GeneratePlan(ByVal pb_AutoGen As Boolean)
  '*** ByVal pb_AutoGen As Boolean '*** (added Jen 2.14)
  
  Dim ws_Cnd As String, wRs_OrdPln As MDORowSet, wi_i As Integer, wi_j As Integer
  Dim wi_Ctr As Integer, wi_Tmp As Integer, wi_Prev As Integer
  Dim ws_PlnStrtDt As String, ws_PlnDt As String
  Dim wsArr_OpPrdStgFlds(0 To 5) As String, wsArr_OpdPrdStgFlds(1 To 5) As String, wsArr_DsgLd(1 To 5) As String
  Dim wfArr_PlnQty(1 To 5) As Single, wfArr_LineCap(1 To 5) As Single, wiArr_GapDys(1 To 5) As Integer
  Dim wRs_LineLd As MDORowSet, wf_BalLd As Single, wb_Flag As Boolean
  Dim wf_BalQty As Single, wf_PlnQty As Single
  Dim wf_MaxQty As Single, wf_MinQty As Single

  Dim ws_TmpTbl As String, wsArr_TmpTblFlds(1 To 5) As String, wRs_TmpTbl As MDORowSet
  Dim wf_CarryFwd As Single, wf_ExistQty As Single
  Dim wi_PlnDays As Integer
  
  '*** (Jen 2.14 Next)
  Dim ws_WDiaAvlblDt As String, ws_HDiaAvlblDt As String
  ws_WDiaAvlblDt = moCn.GetFldVal("Select PMCd from Param where PTyp= 'PRDSTG' and PValue1= 'W'")
  ws_HDiaAvlblDt = moCn.GetFldVal("Select PMCd from Param where PTyp= 'PRDSTG' and PValue1= 'H'")
  '*** (Jen 2.14 Next)
  
  wi_PlnDays = moCn.GetFldVal("Select PNum From Param where PTyp= 'PLNDAYS' and PMCd= 'PLNDAYS'")
  If wi_PlnDays <= 0 Then DispMsg "Number Of Days For Which Plan Has To Be Generated Should Be Specified In Param('PLNDAYS')", etError: Exit Sub
  
  ws_TmpTbl = "#TmpPlan"
  Call CreateTmpTbl(ws_TmpTbl)
  
  ws_Cnd = ADC.RepCond
  ws_Cnd = IIF(ws_Cnd <> "", " and " + ws_Cnd, "")
  '*** If the record exists in the Order Plan Date then Exit Sub ***'
  If moCn.RecSeek("Select 'x' from OrdPlnDt join OrdDsg on OpdCoCd= OdCoCd and " + _
                  "OpdTc= OdTc and OpdYy= OdYy and OpdChr= OdChr and OpdNo= OdNo and " + _
                  "OpdSr= OdSr join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
                  "OmChr= OdChr and OmNo= OdNo where OdCoCd= '" + gs_CoCd + "' " + ws_Cnd) Then
    DispMsg "Select The 'Reset Plan' Option Before 'Generate Plan' As Rec Exists In OrdPlanDate", etError: Exit Sub
  End If
  
  '*** If Plan Date records do not exist then Open Order Plan RecordSet with the Given Scope  ***
  Set wRs_OrdPln = moCn.OpenRes("Select OpTc, OpYy, OpChr, OpNo, OpSr, OpLine, " + _
                   "OpWDiaAvlblDt, OpHDiaAvlblDt, OpWipP1, OpWipP2, OpWipP3, OpWipP4, OpWipP5, " + _
                   "(OdPrdQty- OdFgQty- OpWipP1- OpWipP2- OpWipP3- OpWipP4- OpWipP5) as qFreshPlnQty, OdPrdQty, " + _
                   "IsNull(DL1.DldP1, IsNull(DL2.DldP1, 0)) as DldP1, " + _
                   "IsNull(DL1.DldP2, IsNull(DL2.DldP2, 0)) as DldP2, " + _
                   "IsNull(DL1.DldP3, IsNull(DL2.DldP3, 0)) as DldP3, " + _
                   "IsNull(DL1.DldP4, IsNull(DL2.DldP4, 0)) as DldP4, " + _
                   "IsNull(DL1.DldP5, IsNull(DL2.DldP5, 0)) as DldP5 " + _
                   "from OrdPln join OrdDsg on OdCoCd= OpCoCd and " + _
                   "OdTc= OpTc and OdYy= OpYy and OdChr= OpChr and OdNo= OpNo and OdSr= OpSr " + _
                   "join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr " + _
                   "and OmNo= OdNo join DsgMst on DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
                   "left outer join DsgLoad DL1 on DL1.DldDmCd= OdDmCd and DL1.DldWH= OdWH " + _
                   "left outer join DsgLoad DL2 on DL2.DldPrdCtg= DmPrdCtg and DL2.DldDmCd= '' and DL2.DldWH= OdWH where OpCoCd= '" + gs_CoCd + "' " + ws_Cnd + _
                   " Order By OpLine, OdPrtCd, OdDelDt, OdTc, OdYy, OdChr, OdNo, OdSr ")

  mfArr_AllLineCap(1) = moCn.GetFldVal("Select PcP1 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + ctAllLineCd + "' ")
  mfArr_AllLineCap(2) = moCn.GetFldVal("Select PcP2 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + ctAllLineCd + "' ")
  mfArr_AllLineCap(3) = moCn.GetFldVal("Select PcP3 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + ctAllLineCd + "' ")
  mfArr_AllLineCap(4) = moCn.GetFldVal("Select PcP4 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + ctAllLineCd + "' ")
  mfArr_AllLineCap(5) = moCn.GetFldVal("Select PcP5 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + ctAllLineCd + "' ")

  wsArr_OpPrdStgFlds(1) = "OpWipP1"
  wsArr_OpPrdStgFlds(2) = "OpWipP2"
  wsArr_OpPrdStgFlds(3) = "OpWipP3"
  wsArr_OpPrdStgFlds(4) = "OpWipP4"
  wsArr_OpPrdStgFlds(5) = "OpWipP5"
  wsArr_OpPrdStgFlds(0) = "qFreshPlnQty"

  wsArr_OpdPrdStgFlds(1) = "OpdP1"
  wsArr_OpdPrdStgFlds(2) = "OpdP2"
  wsArr_OpdPrdStgFlds(3) = "OpdP3"
  wsArr_OpdPrdStgFlds(4) = "OpdP4"
  wsArr_OpdPrdStgFlds(5) = "OpdP5"

  wsArr_DsgLd(1) = "DldP1"
  wsArr_DsgLd(2) = "DldP2"
  wsArr_DsgLd(3) = "DldP3"
  wsArr_DsgLd(4) = "DldP4"
  wsArr_DsgLd(5) = "DldP5"

  '*** (Jen 2.14)
  If pb_AutoGen = True Then
    wiArr_GapDys(1) = ADC("wAgGapDysP1")
    wiArr_GapDys(2) = ADC("wAgGapDysP2")
    wiArr_GapDys(3) = ADC("wAgGapDysP3")
    wiArr_GapDys(4) = ADC("wAgGapDysP4")
    wiArr_GapDys(5) = ADC("wAgGapDysP5")
  Else
  '*** (Jen 2.14)
  
    wiArr_GapDys(1) = ADC("wGpGapDysP1")
    wiArr_GapDys(2) = ADC("wGpGapDysP2")
    wiArr_GapDys(3) = ADC("wGpGapDysP3")
    wiArr_GapDys(4) = ADC("wGpGapDysP4")
    wiArr_GapDys(5) = ADC("wGpGapDysP5")
    
  '*** (Jen 2.14)
  End If
  '*** (Jen 2.14)

  wsArr_TmpTblFlds(1) = "TpP1"
  wsArr_TmpTblFlds(2) = "TpP2"
  wsArr_TmpTblFlds(3) = "TpP3"
  wsArr_TmpTblFlds(4) = "TpP4"
  wsArr_TmpTblFlds(5) = "TpP5"

  With wRs_OrdPln
    Do While Not (.EOF Or .BOF)
      ws_PlnStrtDt = NextDate(0, moCn.SrvrDate)
      '*** Bef 09/01/04 ws_PlnDt = ws_PlnStrtDt
      '*** Get the line Capacity in an Array ***'
      wfArr_LineCap(1) = moCn.GetFldVal("Select PcP1 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + wRs_OrdPln!OpLine + "' ")
      wfArr_LineCap(2) = moCn.GetFldVal("Select PcP2 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + wRs_OrdPln!OpLine + "' ")
      wfArr_LineCap(3) = moCn.GetFldVal("Select PcP3 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + wRs_OrdPln!OpLine + "' ")
      wfArr_LineCap(4) = moCn.GetFldVal("Select PcP4 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + wRs_OrdPln!OpLine + "' ")
      wfArr_LineCap(5) = moCn.GetFldVal("Select PcP5 from PpcCap where PcCoCd= '" + gs_CoCd + "' and PcLine= '" + wRs_OrdPln!OpLine + "' ")
      '*** For all the Five Production Stages and the Fresh Plan Quantity ***'
      
      For wi_i = 5 To 0 Step -1
        If wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) > 0 Then ws_PlnDt = ws_PlnStrtDt '*** 08/01/04
      
        wi_Ctr = IIF(wi_i >= 1, wi_i, 0)    '*** 06/01/04
        wi_j = wi_Ctr + 1
        
        '*** Get the Minimum and Maximum Quantity From the Head ***'
        wf_MinQty = IIF(wi_i = 0, 9999999, 0)   '*** 08/01/04
        wf_MaxQty = 9999999
        
        '*** Take Minimum of (Ord Qty * Min Ord Percent /100) or Min Order Qty ***'
        '*** Bef 08/01/04 If wf_MinQty > wRs_OrdPln!OdPrdQty * (mf_hMinOrdPrcnt / 100) Then wf_MinQty = wRs_OrdPln!OdPrdQty * (mf_hMinOrdPrcnt / 100)
        If wf_MinQty > wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) * (mf_hMinOrdPrcnt / 100) Then wf_MinQty = wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) * (mf_hMinOrdPrcnt / 100) '*** 08/01/04
        If wf_MinQty > mf_hMinOrdQty Then wf_MinQty = mf_hMinOrdQty
        '*** Take Maximum as Max Ord Qty
        If wf_MaxQty > mf_hMaxOrdQty Then wf_MaxQty = mf_hMaxOrdQty
        
        '*** If Any Qty is present in the Production Stage then Go into the If
        '    Else Start with the Next Production Stage ***'
        If wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) > 0 Then
          wi_Prev = wi_Ctr
          '*** Loop Till The Qty is all planned for ***'
          Do While wi_j > wi_Ctr And wi_j <= 5
            If wi_j > 1 Then wi_Prev = wi_j - 1
            '*** Loop if Line Capacity of the Prd Stg is Zero or Design Load is Zero then ***'
            ' **** wi_j is incremented so as to start with the Next Production Stage *****
            Do While wfArr_LineCap(wi_j) = 0 Or wRs_OrdPln(wsArr_DsgLd(wi_j)) = 0
              wi_j = wi_j + 1
              If wi_j >= 5 Then Exit Do
            Loop
            If wi_j > 5 Then Exit Do
            '*** For the Production Stages after P1  ***'
            If wi_Prev = wi_Ctr Then
              moCn.Execute ("delete from " + ws_TmpTbl)
              '**** For the Bags in the Production Stages (Already existing Orders) *****
              wf_BalQty = wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i))
              
              '****** Loop till the Balanced Quantity is Higher than Zero and the Minimum Qty *****
              '$$$
              '*** Bef 09/01/04 Do While Round(wf_BalQty, 1) > 0 And _
                ((wf_BalQty = wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) And Round(wf_BalQty, 1) >= Round(wf_MinQty, 1)) _
                  Or (wf_BalQty < wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i))))
              '*** 09/01/04
              Do While Round(wf_BalQty, 1) > 0
              '*** 09/01/04
              '$$$
                
                '**** wb_Flag Stores the Status whether Some Quantity is Stil Balanced to Plan for ***'
                wb_Flag = False
                Do While wb_Flag = False ' *** Loop till Quantity is Balanced to Plan For
                  '*** Plan Qty is either Balance Quantity if that is less or The Maximum Order Quantity ****'
                  '*** Get the Diamond Availability Dates for P1 or P4 If Plan Date is Greater Consider Plan Date else Consider Availability Date as the Plan Date for that Process ***'
                  '@@@
                  
                  '*** (Jen 2.14 Next)
                  If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("Opd" + ws_WDiaAvlblDt) Then
                    ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpWDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpWDiaAvlblDt)
                  End If
                  If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("Opd" + ws_HDiaAvlblDt) Then
                    ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpHDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpHDiaAvlblDt)
                  End If
                  '*** (Jen 2.14 Next)
                  
                  '*** (Bef 2.14 Next)
                  'If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("OpdP1") Then
                  '  ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpWDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpWDiaAvlblDt)
                  'End If
                  'If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("OpdP4") Then
                  '  ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpHDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpHDiaAvlblDt)
                  'End If
                  '*** (Bef 2.14 Next)
                  '@@@
                  '*** Get the Line Load ***'
                  '6.1
                  Set wRs_LineLd = moCn.OpenRes("Set dateformat " + DtFmtStr() + " " + _
                      "Select IsNull(sum(" + wsArr_OpdPrdStgFlds(wi_j) + "* IsNull(DL1." + wsArr_DsgLd(wi_j) + ", IsNull(DL2." + wsArr_DsgLd(wi_j) + ", 0))* (case when OpLine= '" + wRs_OrdPln!OpLine + "' then 1 else 0 end)), 0) as qLineLd, " + _
                      "IsNull(sum(" + wsArr_OpdPrdStgFlds(wi_j) + "* IsNull(DL1." + wsArr_DsgLd(wi_j) + ", IsNull(DL2." + wsArr_DsgLd(wi_j) + ", 0))), 0) as qAllLineLd " + _
                      "from OrdPlnDt join OrdPln on OpCoCd= OpdCoCd and OpTc= OpdTc and " + _
                      "OpYy= OpdYy and OpChr= OpdChr and OpNo= OpdNo and OpSr= OpdSr " + _
                      "join OrdDsg on OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
                      "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr " + _
                      "join DsgMst on DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
                      "left outer join DsgLoad DL1 on DL1.DldDmCd= OdDmCd and DL1.DldWH= OdWH " + _
                      "left outer join DsgLoad DL2 on DL2.DldPrdCtg= DmPrdCtg and DL2.DldDmCd= '' and DL2.DldWH= OdWH " + _
                      "Where OpCoCd= '" + gs_CoCd + "' and OpdPlnDt= '" + ws_PlnDt + "' ")
    
                  wf_BalLd = 0
                  If Not (wRs_LineLd.EOF Or wRs_LineLd.BOF) Then
                  '*** Balance Load = Line Capacity For the Prd Stg * Max Line Utilization /100 - (Prd Stg Qty * Dsg Load)
                    If wfArr_LineCap(wi_j) > 0 And wRs_OrdPln(wsArr_DsgLd(wi_j)) > 0 Then wf_BalLd = wfArr_LineCap(wi_j) * (mf_hMaxLineUtil / 100) - wRs_LineLd!qLineLd
                  ' **** if balance load is more than the All Line Capacity then Bal Load = All Line Capacity - All Line Load
                    If wf_BalLd > mfArr_AllLineCap(wi_j) - wRs_LineLd!qAllLineLd Then wf_BalLd = mfArr_AllLineCap(wi_j) - wRs_LineLd!qAllLineLd
                  End If
                  
                  '*** 09/01/04
                  If wf_PlnQty = 0 Then
                    wf_PlnQty = Round(IIF(wf_BalQty > wf_MaxQty, wf_MaxQty, wf_BalQty), 1)
                    wf_ExistQty = moCn.GetFldVal("Select " + wsArr_OpdPrdStgFlds(wi_j) + " from OrdPlnDt " + _
                         "where OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                         "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                         "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                         "and OpdPlnDt= '" + ws_PlnDt + "' ")
                    If wf_BalLd > 0 And wf_BalLd < wf_PlnQty * wRs_OrdPln(wsArr_DsgLd(wi_j)) Then
                      If Int(Round(MWLib.Div(wf_BalLd, wRs_OrdPln(wsArr_DsgLd(wi_j))), 4)) + _
                         wf_ExistQty >= wf_MinQty Then
                        wf_PlnQty = Int(Round(MWLib.Div(wf_BalLd, wRs_OrdPln(wsArr_DsgLd(wi_j))), 4))
                      Else
                        wf_PlnQty = 0
                      End If
                    ElseIf wf_BalLd = 0 Then
                      wf_PlnQty = 0
                    End If
                    wf_BalQty = wf_BalQty - wf_PlnQty
                    
                    If wf_PlnQty > 0 And wf_PlnQty + wf_ExistQty > wf_MaxQty Then
                      wf_BalQty = wf_BalQty + wf_PlnQty + wf_ExistQty - wf_MaxQty
                      wf_PlnQty = wf_MaxQty - wf_ExistQty
                    End If
                  End If
                  '*** 09/01/04
                  
                  '**** if Balance Load is higher then Plan Qty
                  If wf_PlnQty > 0 And wf_BalLd >= wf_PlnQty * wRs_OrdPln(wsArr_DsgLd(wi_j)) Then
                    If DateVal(ws_PlnDt) > DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then GoTo XNext     '*** 06/01/04
                    '*** 08/01/04
                    For wi_Tmp = 1 To 5
                      wfArr_PlnQty(wi_Tmp) = IIF(wi_Tmp = wi_j, wf_PlnQty, 0)
                    Next wi_Tmp
                    '*** 08/01/04
                    
                    If Not moCn.RecSeek("Select 'x' from " + ws_TmpTbl + " where TpPlnDt= '" + ws_PlnDt + "' ") And DateVal(ws_PlnDt) <= DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then
                      '6.1
                      moCn.Execute ("Set DateFormat " + DtFmtStr() + " " + _
                                    "Insert Into " + ws_TmpTbl + " (TpPlnDt, TpP1, TpP2, " + _
                                    "TpP3, TpP4, TpP5) " + _
                                    "values('" + ws_PlnDt + "' , " + CStr(wfArr_PlnQty(1)) + " " + _
                                    ", " + CStr(wfArr_PlnQty(2)) + ", " + CStr(wfArr_PlnQty(3)) + " " + _
                                    ", " + CStr(wfArr_PlnQty(4)) + ", " + CStr(wfArr_PlnQty(5)) + ")")
                    Else  '6.1
                      moCn.Execute ("set dateformat " + DtFmtStr() + " " + _
                                    "update " + ws_TmpTbl + " set " + wsArr_TmpTblFlds(wi_j) + "= " + wsArr_TmpTblFlds(wi_j) + " + " + CStr(wf_PlnQty) + " where " + _
                                    "TpPlnDt= '" + ws_PlnDt + "' ")
                    End If
                    
                    If Not moCn.RecSeek("Select 'x' from OrdPlnDt where OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                                    "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                                    "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                                    "and OpdPlnDt= '" + ws_PlnDt + "' ") Then
                      If wfArr_PlnQty(1) > 0 Or wfArr_PlnQty(2) > 0 Or wfArr_PlnQty(3) > 0 Or _
                         wfArr_PlnQty(4) > 0 Or wfArr_PlnQty(5) > 0 Then
                        '6.1
                        moCn.Execute ("Set DateFormat " + DtFmtStr() + " " + _
                                      "Insert Into OrdPlnDt (OpdCoCd, OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, " + _
                                      "OpdPlnDt, OpdP1, OpdP2, OpdP3, OpdP4, OpdP5, ModUsr, ModDt, ModTime) " + _
                                      "values('" + gs_CoCd + "', '" + !OpTc + "', '" + !OpYy + "', '" + _
                                      !OpChr + "', " + CStr(!OpNo) + ", " + CStr(!OpSr) + ", '" + ws_PlnDt + "' " + _
                                      ", " + CStr(wfArr_PlnQty(1)) + ", " + CStr(wfArr_PlnQty(2)) + ", " + CStr(wfArr_PlnQty(3)) + " " + _
                                      ", " + CStr(wfArr_PlnQty(4)) + ", " + CStr(wfArr_PlnQty(5)) + ", '" + ADC.UsrCd + "', " + _
                                      "'" + CStr(moCn.SrvrDate) + "', " + CStr(moCn.SrvrTime) + ")")
                      End If
                      wb_Flag = True
                      wf_PlnQty = 0 '$$$
                    Else  '6.1
                      moCn.Execute ("set dateformat " + DtFmtStr() + " " + _
                                    "update OrdPlnDt set " + wsArr_OpdPrdStgFlds(wi_j) + "= " + wsArr_OpdPrdStgFlds(wi_j) + " + " + CStr(wf_PlnQty) + " where " + _
                                    "OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                                    "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                                    "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                                    "and OpdPlnDt= '" + ws_PlnDt + "' ")
                      wb_Flag = True
                      wf_PlnQty = 0 '$$$
                    End If
                  End If
                  '*** If Bal Load is Less than the Plan Qty then  Skip Day and Check for the Next Day
                  If Round(wf_BalQty, 1) > 0 Or wb_Flag = False Then ws_PlnDt = NextDate(1, ws_PlnDt)
                  If DateVal(ws_PlnDt) > DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then GoTo XNext
                Loop
              Loop
           
            '###%%%
            Else  '6.1
              Set wRs_TmpTbl = moCn.OpenRes("Set dateformat " + DtFmtStr() + " " + _
                               "Select * from " + ws_TmpTbl + " Order By TpPlnDt")
              wf_CarryFwd = 0
              Do While Not (wRs_TmpTbl.EOF Or wRs_TmpTbl.BOF)
                '*** 08/01/04
                wf_BalQty = wRs_TmpTbl(wsArr_TmpTblFlds(wi_Prev)) + wf_CarryFwd
                wf_CarryFwd = 0
                ws_PlnDt = ""
                '*** 08/01/04
                wf_PlnQty = 0
                '*** Bef 09/01/04 Do While Round(wf_BalQty, 1) > 0 And _
                  ((wf_BalQty = wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i)) And Round(wf_BalQty, 1) >= Round(wf_MinQty, 1)) _
                    Or (wf_BalQty < wRs_OrdPln(wsArr_OpPrdStgFlds(wi_i))))
                '*** 09/01/04
                Do While Round(wf_BalQty, 1) > 0
                '*** 09/01/04
                
                  If ws_PlnDt = "" Then ws_PlnDt = "01/01/80"
                  ws_PlnDt = IIF(DateVal(ws_PlnDt) >= DateVal(NextDate(wiArr_GapDys(wi_Prev), wRs_TmpTbl!TpPlnDt)), ws_PlnDt, NextDate(wiArr_GapDys(wi_Prev), wRs_TmpTbl!TpPlnDt))
                  
                  '@@@
                  '*** (Jen 2.14 Next)
                  If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("Opd" + ws_WDiaAvlblDt) Then
                    ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpWDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpWDiaAvlblDt)
                  End If
                  If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("Opd" + ws_HDiaAvlblDt) Then
                    ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpHDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpHDiaAvlblDt)
                  End If
                  '*** (Jen 2.14 Next)
                  
                  '*** (Bef 2.14 Next)
                  'If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("OpdP1") Then
                  '  ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpWDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpWDiaAvlblDt)
                  'End If
                  'If UCase(wsArr_OpdPrdStgFlds(wi_j)) = UCase("OpdP4") Then
                  '  ws_PlnDt = IIF(DateVal(ws_PlnDt) > DateVal(wRs_OrdPln!OpHDiaAvlblDt), ws_PlnDt, wRs_OrdPln!OpHDiaAvlblDt)
                  'End If
                  '*** (Bef 2.14 Next)
                  '@@@
                  
                  wb_Flag = False
                  Do While wb_Flag = False
                  '6.1
                    Set wRs_LineLd = moCn.OpenRes("Set dateformat " + DtFmtStr() + " " + _
                       "Select IsNull(sum(" + wsArr_OpdPrdStgFlds(wi_j) + "* IsNull(DL1." + wsArr_DsgLd(wi_j) + ", IsNull(DL2." + wsArr_DsgLd(wi_j) + ", 0))* (case when OpLine= '" + wRs_OrdPln!OpLine + "' then 1 else 0 end)), 0) as qLineLd, " + _
                       "IsNull(sum(" + wsArr_OpdPrdStgFlds(wi_j) + "* IsNull(DL1." + wsArr_DsgLd(wi_j) + ", IsNull(DL2." + wsArr_DsgLd(wi_j) + ", 0))), 0) as qAllLineLd " + _
                       "from OrdPlnDt join OrdPln on OpCoCd= OpdCoCd and OpTc= OpdTc and " + _
                       "OpYy= OpdYy and OpChr= OpdChr and OpNo= OpdNo and OpSr= OpdSr " + _
                       "join OrdDsg on OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and " + _
                       "OdChr= OpChr and OdNo= OpNo and OdSr= OpSr " + _
                       "join DsgMst on DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
                       "left outer join DsgLoad DL1 on DL1.DldDmCd= OdDmCd and DL1.DldWH= OdWH " + _
                       "left outer join DsgLoad DL2 on DL2.DldPrdCtg= DmPrdCtg and DL2.DldDmCd= '' and DL2.DldWH= OdWH " + _
                       "Where OpCoCd= '" + gs_CoCd + "' and OpdPlnDt= '" + ws_PlnDt + "' ")
      
                    wf_BalLd = 0
                    If Not (wRs_LineLd.EOF Or wRs_LineLd.BOF) Then
                      If wfArr_LineCap(wi_j) > 0 And wRs_OrdPln(wsArr_DsgLd(wi_j)) > 0 Then wf_BalLd = wfArr_LineCap(wi_j) * (mf_hMaxLineUtil / 100) - wRs_LineLd!qLineLd
                      If wf_BalLd > mfArr_AllLineCap(wi_j) - wRs_LineLd!qAllLineLd Then wf_BalLd = mfArr_AllLineCap(wi_j) - wRs_LineLd!qAllLineLd
                    End If
      
                    '*** 08/01/04
                    If wf_PlnQty = 0 Then
                      wf_PlnQty = Round(IIF(wf_BalQty > wf_MaxQty, wf_MaxQty, wf_BalQty), 1)
                      wf_ExistQty = moCn.GetFldVal("Select " + wsArr_OpdPrdStgFlds(wi_j) + " from OrdPlnDt " + _
                           "where OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                           "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                           "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                           "and OpdPlnDt= '" + ws_PlnDt + "' ")
                      If wf_BalLd > 0 And wf_BalLd < wf_PlnQty * wRs_OrdPln(wsArr_DsgLd(wi_j)) Then
                        If Int(Round(MWLib.Div(wf_BalLd, wRs_OrdPln(wsArr_DsgLd(wi_j))), 4)) + _
                           wf_ExistQty >= wf_MinQty Then
                          wf_PlnQty = Int(Round(MWLib.Div(wf_BalLd, wRs_OrdPln(wsArr_DsgLd(wi_j))), 4))
                        Else
                          wf_PlnQty = 0
                        End If
                      ElseIf wf_BalLd = 0 Then
                        wf_PlnQty = 0
                      End If
                      If wf_PlnQty > 0 And wf_BalQty = wf_PlnQty And wf_PlnQty + _
                         wf_ExistQty < wf_MinQty And _
                         moCn.RecSeek("Select 'x' from " + ws_TmpTbl + " where  " + _
                              "TpPlnDt> '" + CStr(wRs_TmpTbl!TpPlnDt) + "' and " + wsArr_TmpTblFlds(wi_Prev) + "> 0 ") Then
                        wf_CarryFwd = wf_BalQty: wf_BalQty = 0: wf_PlnQty = 0: wb_Flag = True
                      Else
                        wf_BalQty = wf_BalQty - wf_PlnQty
                      End If
                      '*** 09/01/04
                      If wf_PlnQty > 0 And wf_PlnQty + wf_ExistQty > wf_MaxQty Then
                        wf_BalQty = wf_BalQty + wf_PlnQty + wf_ExistQty - wf_MaxQty
                        wf_PlnQty = wf_MaxQty - wf_ExistQty
                      End If
                      '*** 09/01/04
                    End If
                    '*** 08/01/04
            
                    If wf_PlnQty > 0 And wf_BalLd >= wf_PlnQty * wRs_OrdPln(wsArr_DsgLd(wi_j)) Then
                      If DateVal(ws_PlnDt) > DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then GoTo XNext   '*** 06/01/04
                      '*** 08/01/04
                      For wi_Tmp = 1 To 5
                        wfArr_PlnQty(wi_Tmp) = IIF(wi_Tmp = wi_j, wf_PlnQty, 0)
                      Next wi_Tmp
                      '*** 08/01/04
                      '6.1 date format added in if and else part
                      If Not moCn.RecSeek("Select 'x' from " + ws_TmpTbl + " where TpPlnDt= '" + ws_PlnDt + "' ") And DateVal(ws_PlnDt) <= DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then
                        moCn.Execute ("Set DateFormat " + DtFmtStr() + " " + _
                                      "Insert Into " + ws_TmpTbl + " (TpPlnDt, TpP1, TpP2, " + _
                                      "TpP3, TpP4, TpP5) " + _
                                      "values('" + ws_PlnDt + "' , " + CStr(wfArr_PlnQty(1)) + " " + _
                                      ", " + CStr(wfArr_PlnQty(2)) + ", " + CStr(wfArr_PlnQty(3)) + " " + _
                                      ", " + CStr(wfArr_PlnQty(4)) + ", " + CStr(wfArr_PlnQty(5)) + ")")
                      Else
                        moCn.Execute ("set dateformat " + DtFmtStr() + " " + _
                                      "update " + ws_TmpTbl + " set " + wsArr_TmpTblFlds(wi_j) + "= " + wsArr_TmpTblFlds(wi_j) + " + " + CStr(wf_PlnQty) + " where " + _
                                      "TpPlnDt= '" + ws_PlnDt + "' ")
                      End If
                      
                      If Not moCn.RecSeek("Select 'x' from OrdPlnDt where OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                                      "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                                      "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                                      "and OpdPlnDt= '" + ws_PlnDt + "' ") Then
                          
                        If wfArr_PlnQty(1) > 0 Or wfArr_PlnQty(2) > 0 Or wfArr_PlnQty(3) > 0 Or _
                           wfArr_PlnQty(4) > 0 Or wfArr_PlnQty(5) > 0 Then
                          '6.1 date format added in if and else
                          moCn.Execute ("Set DateFormat " + DtFmtStr() + " " + _
                                        "Insert Into OrdPlnDt (OpdCoCd, OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, " + _
                                        "OpdPlnDt, OpdP1, OpdP2, OpdP3, OpdP4, OpdP5, ModUsr, ModDt, ModTime) " + _
                                        "values('" + gs_CoCd + "', '" + !OpTc + "', '" + !OpYy + "', '" + _
                                        !OpChr + "', " + CStr(!OpNo) + ", " + CStr(!OpSr) + ", '" + ws_PlnDt + "' " + _
                                        ", " + CStr(wfArr_PlnQty(1)) + ", " + CStr(wfArr_PlnQty(2)) + ", " + CStr(wfArr_PlnQty(3)) + " " + _
                                        ", " + CStr(wfArr_PlnQty(4)) + ", " + CStr(wfArr_PlnQty(5)) + ", '" + ADC.UsrCd + "', " + _
                                        "'" + CStr(moCn.SrvrDate) + "', " + CStr(moCn.SrvrTime) + ")")
                        End If
                        wb_Flag = True
                        wf_PlnQty = 0
                      Else
                        moCn.Execute ("set dateformat " + DtFmtStr() + " " + _
                                      "update OrdPlnDt set " + wsArr_OpdPrdStgFlds(wi_j) + "= " + wsArr_OpdPrdStgFlds(wi_j) + " + " + CStr(wf_PlnQty) + " where " + _
                                      "OpdCoCd= '" + gs_CoCd + "' and OpdTc= '" + wRs_OrdPln!OpTc + "' " + _
                                      "and OpdYy= '" + wRs_OrdPln!OpYy + "' and OpdChr= '" + wRs_OrdPln!OpChr + "' " + _
                                      "and OpdNo= " + CStr(wRs_OrdPln!OpNo) + " and OpdSr= " + CStr(wRs_OrdPln!OpSr) + " " + _
                                      "and OpdPlnDt= '" + ws_PlnDt + "' ")
                        wb_Flag = True
                        wf_PlnQty = 0
                      End If
                    End If
                    '$$$
                    If Round(wf_BalQty, 1) > 0 Or wb_Flag = False Then ws_PlnDt = NextDate(1, ws_PlnDt)
                    '$$$
                    If DateVal(ws_PlnDt) > DateVal(NextDate(wi_PlnDays, moCn.SrvrDate)) Then GoTo XNext
                  Loop
                Loop
                wRs_TmpTbl.MoveNext
              Loop
              
            '###%%%
            End If
  
XNext:
            wi_j = wi_j + 1
          Loop
        End If
      Next wi_i
      .MoveNext
    Loop
  End With

  If GrdOp.Row <= 0 And GrdOp.Rows - 1 > 0 Then GrdOp.Row = 1
  If GrdOp.Rows - 1 > 0 And GrdOp.Row > 0 Then
    GrdOpd.Init
    GrdOpd.RefreshData
  End If
  
  Set wRs_OrdPln = Nothing
  Set wRs_LineLd = Nothing
End Sub
Private Sub CreateTmpTbl(ByVal ms_Tbl As String)
  Dim ws_str
  ws_str = " CREATE TABLE [dbo].[" + ms_Tbl + "] ( " + _
           "[TpPlnDt] [smalldatetime] not null, " + _
           "[TpP1] [float] not null, " + _
           "[TpP2] [float] not null, " + _
           "[TpP3] [float] not null, " + _
           "[TpP4] [float] not null, " + _
           "[TpP5] [float] not null " + _
           ") ON [PRIMARY] "
  
  'moCn.Execute (ws_Str)

  '*** Creating Temporary Table ***
  moCn.CreateTmpTable ms_Tbl, ws_str
End Sub
Private Sub PlanRedo()
  Dim ws_FldName As String, wf_Qty As Single, ws_Dt As String, wsArr_OpdPrdStgFlds(1 To 5) As String
  Dim wiArr_GapDays(1 To 5) As Integer, ws_PrevDt As String
  Dim wi_i As Integer, wi_Start As Integer, wi_j As Integer, wi_Row As Integer

  If GrdOpd.Row = 0 Then Exit Sub
  If GrdOpd.Col <= 1 Or GrdOpd.Col >= 6 Then Exit Sub

  ws_FldName = GrdOpd.ColProp(GrdOpd.Col).ColName
  wf_Qty = GrdOpd.Value(GrdOpd.Row, ws_FldName)
  ws_Dt = GrdOpd.Value(GrdOpd.Row, "OpdPlnDt")
  ws_PrevDt = ws_Dt
  If wf_Qty = 0 Then Exit Sub

  wsArr_OpdPrdStgFlds(1) = "OpdP1"
  wsArr_OpdPrdStgFlds(2) = "OpdP2"
  wsArr_OpdPrdStgFlds(3) = "OpdP3"
  wsArr_OpdPrdStgFlds(4) = "OpdP4"
  wsArr_OpdPrdStgFlds(5) = "OpdP5"

  For wi_i = GrdOpd.Rows - 1 To GrdOpd.Row + 1 Step -1
    GrdOpd.Del_Rec False, wi_i
  Next wi_i

  For wi_i = 1 To 5
    wiArr_GapDays(wi_i) = moCn.GetFldVal("Select PNum1 from Param where PTyp= 'PRDSTG' and Str(PNum, 2, 0)= Str(" + CStr(wi_i) + ", 2, 0) ")
    If UCase(wsArr_OpdPrdStgFlds(wi_i)) = UCase(ws_FldName) Then wi_Start = wi_i + 1
  Next wi_i
  
  With GrdOpd
    .StartCopy
    For wi_j = wi_Start To 5
      ws_Dt = NextDate(wiArr_GapDays(wi_j - 1), ws_Dt)
      If ws_PrevDt <> ws_Dt Then
        .AddItem
        wi_Row = .Rows - 1
        .Value(wi_Row, "OpdPlnDt") = ws_Dt
        For wi_i = 1 To 5
          .Value(wi_Row, wsArr_OpdPrdStgFlds(wi_i)) = 0
        Next wi_i
        .Value(wi_Row, wsArr_OpdPrdStgFlds(wi_j)) = wf_Qty
        .SaveRec (wi_Row)
      Else
        If wi_Row = 0 Then wi_Row = .Rows - 1
        .Value(wi_Row, wsArr_OpdPrdStgFlds(wi_j)) = wf_Qty
        .SaveRec (wi_Row)
      End If
      ws_PrevDt = ws_Dt
    Next wi_j
    .EndCopy
  End With
End Sub
Private Sub DesignDetails()
  '*** Procedure to get the Order Design Details in the Design Details Frame
  If GrdOp.Rows - 1 = 0 Then Exit Sub
  Dim wf_OpRow As Single, wRs_OrdDsg As MwfLib.MDORowSet
  Dim ws_3DPth As String, ws_HskPth  As String
  
  wf_OpRow = IIF(GrdOp.Row > 0, GrdOp.Row, 1)
  
  Set wRs_OrdDsg = moCn.OpenRes("Select OdDmSz, OdSfx, OdKt, DmCtg, OdVaCtg " + _
                   "from OrdDsg join DsgMst on DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
                   "Where OdCoCd='" + gs_CoCd + "' and " + _
                   "OdTc= '" + GrdOp.Value(wf_OpRow, "OpTc") + "' and " + _
                   "OdYy= '" + GrdOp.Value(wf_OpRow, "OpYy") + "' and " + _
                   "OdChr= '" + GrdOp.Value(wf_OpRow, "OpChr") + "' and " + _
                   "OdNo= " + CStr(GrdOp.Value(wf_OpRow, "OpNo")) + " and " + _
                   "OdSr= " + CStr(GrdOp.Value(wf_OpRow, "OpSr")))
  
  If Not (wRs_OrdDsg.EOF Or wRs_OrdDsg.BOF) Then
    ADC("wDdDmSz") = wRs_OrdDsg!OdDmSz
    ADC("wDdSfx") = wRs_OrdDsg!OdSfx
    ADC("wDdKt") = wRs_OrdDsg!OdKt
    ADC("wDdDmSz") = wRs_OrdDsg!OdDmSz
    ADC("wDdDmCtg") = wRs_OrdDsg!DmCtg
    ADC("wDdVaCtg") = wRs_OrdDsg!OdVaCtg
  End If
  
  FraPic3D.Left = md_FraPic3DLft: FraPic3D.Top = md_FraPic3DTop
  FraPic3D.Height = mf_FraPic3DHt: FraPic3D.Width = mf_FraPic3DWdt
  FraPicHsk.Left = md_FraPicHskLft: FraPicHsk.Top = md_FraPicHskTop
  FraPicHsk.Height = mf_FraPicHskHt: FraPicHsk.Width = mf_FraPicHskWdt
  
  ws_3DPth = GetPictPath("DM", False, GrdOp.Value(wf_OpRow, "wOdDmCd"))
  ws_HskPth = GetPictPath("DM", True, GrdOp.Value(wf_OpRow, "wOdDmCd"))
  
  Call ShowPic(md_Pic3DHt, md_Pic3DWdt, ws_3DPth, Pic3D)
  Call ShowPic(md_PicHskHt, md_PicHskWdt, ws_HskPth, PicHsk)
  
  Set wRs_OrdDsg = Nothing
End Sub
Private Sub ShowLoad()
  '*** Procedure to Show Load Utilised
  Dim wRs_ShwLd As MwfLib.MDORowSet, wi_Row As Integer, ws_CommStr As String, ws_Cnd As String
  
  If ADC("wSlLine") = "" Then DispMsg "Line Code Should Be Entered", etError: Exit Sub
  If ADC("wSlPlnDtFr") = "01/01/80" Then DispMsg "From Plan Date Cannot Be Blank", etError: Exit Sub
  If ADC("wSlPlnDtFr") > ADC("wSlPlnDtTo") Then DispMsg "To Plan Date Should Be Greater Than From Date", etError: Exit Sub
'6.1
  ws_CommStr = "Set dateformat " + DtFmtStr() + " " + _
               "Select DtDt as qPlnDt, " + _
               "sum(IsNull(OpdP1, 0)* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP1Qty, " + _
               "sum(IsNull(OpdP2, 0)* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP2Qty, " + _
               "sum(IsNull(OpdP3, 0)* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP3Qty, " + _
               "sum(IsNull(OpdP4, 0)* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP4Qty, " + _
               "sum(IsNull(OpdP5, 0)* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP5Qty, " + _
               "sum(IsNull(OpdP1, 0)* IsNull(DL1.DldP1, IsNull(DL2.DldP1, 0))* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP1Ld, " + _
               "sum(IsNull(OpdP2, 0)* IsNull(DL1.DldP2, IsNull(DL2.DldP2, 0))* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP2Ld, " + _
               "sum(IsNull(OpdP3, 0)* IsNull(DL1.DldP3, IsNull(DL2.DldP3, 0))* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP3Ld, " + _
               "sum(IsNull(OpdP4, 0)* IsNull(DL1.DldP4, IsNull(DL2.DldP4, 0))* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP4Ld, " + _
               "sum(IsNull(OpdP5, 0)* IsNull(DL1.DldP5, IsNull(DL2.DldP5, 0))* (case when IsNull(OpLine, '')= '" + ADC("wSlLine") + "' then 1 else 0 end)) as qOpdP5Ld, " + _
               "sum(IsNull(OpdP1, 0)* IsNull(DL1.DldP1, IsNull(DL2.DldP1, 0))) as qOpdP1LdAll, " + _
               "sum(IsNull(OpdP2, 0)* IsNull(DL1.DldP2, IsNull(DL2.DldP2, 0))) as qOpdP2LdAll, " + _
               "sum(IsNull(OpdP3, 0)* IsNull(DL1.DldP3, IsNull(DL2.DldP3, 0))) as qOpdP3LdAll, " + _
               "sum(IsNull(OpdP4, 0)* IsNull(DL1.DldP4, IsNull(DL2.DldP4, 0))) as qOpdP4LdAll, " + _
               "sum(IsNull(OpdP5, 0)* IsNull(DL1.DldP5, IsNull(DL2.DldP5, 0))) as qOpdP5LdAll "
               
  ws_CommStr = ws_CommStr + " from DtTbl left outer join OrdPlnDt on OpdPlnDt= DtDt " + _
               "left outer join OrdPln on OpCoCd= OpdCoCd and OpTc= OpdTc and " + _
               "OpYy= OpdYy and OpChr= OpdChr and OpNo= OpdNo and OpSr= OpdSr " + _
               "left outer join OrdDsg on OdCoCd= OpCoCd and OdTc= OpTc and " + _
               "OdYy= OpYy and OdChr= OpChr and OdNo= OpNo and OdSr= OpSr " + _
               "left outer join DsgMst on DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '' " + _
               "left outer join DsgLoad DL1 on DL1.DldDmCd= OdDmCd and DL1.DldWH= OdWH " + _
               "left outer join DsgLoad DL2 on DL2.DldPrdCtg= DmPrdCtg and DL2.DldDmCd= '' and DL2.DldWH= OdWH " + _
               "Where DtTag= 'HOL' and IsNull(OpCoCd, '" + gs_CoCd + "')= '" + gs_CoCd + "' " + _
               "and DtDt>= '" + CStr(ADC("wSlPlnDtFr")) + "' and " + _
               "DtDt<= '" + CStr(ADC("wSlPlnDtTo")) + "' "
               'and DtHolYN = 'N'
  Set wRs_ShwLd = moCn.OpenRes(ws_CommStr + " Group By DtDt Order By convert(varchar(10), DtDt, 112)")
  
  With GrdSL
    .Rows = 1
    .AllowAdd = True
    Do While Not (wRs_ShwLd.EOF Or wRs_ShwLd.BOF)
      .AddItem
      wi_Row = .Rows - 1
      .Value(wi_Row, "wSlDt") = wRs_ShwLd!qPlnDt
      .Value(wi_Row, "wSlP1") = wRs_ShwLd!qOpdP1Qty
      .Value(wi_Row, "wSlP2") = wRs_ShwLd!qOpdP2Qty
      .Value(wi_Row, "wSlP3") = wRs_ShwLd!qOpdP3Qty
      .Value(wi_Row, "wSlP4") = wRs_ShwLd!qOpdP4Qty
      .Value(wi_Row, "wSlP5") = wRs_ShwLd!qOpdP5Qty
      
      .Value(wi_Row, "wSlP1Ld") = wRs_ShwLd!qOpdP1Ld
      .Value(wi_Row, "wSlP2Ld") = wRs_ShwLd!qOpdP2Ld
      .Value(wi_Row, "wSlP3Ld") = wRs_ShwLd!qOpdP3Ld
      .Value(wi_Row, "wSlP4Ld") = wRs_ShwLd!qOpdP4Ld
      .Value(wi_Row, "wSlP5Ld") = wRs_ShwLd!qOpdP5Ld
      
      .Value(wi_Row, "wSlP1LdAll") = wRs_ShwLd!qOpdP1LdAll
      .Value(wi_Row, "wSlP2LdAll") = wRs_ShwLd!qOpdP2LdAll
      .Value(wi_Row, "wSlP3LdAll") = wRs_ShwLd!qOpdP3LdAll
      .Value(wi_Row, "wSlP4LdAll") = wRs_ShwLd!qOpdP4LdAll
      .Value(wi_Row, "wSlP5LdAll") = wRs_ShwLd!qOpdP5LdAll
      
      wRs_ShwLd.MoveNext
    Loop
    .AllowAdd = False
  End With
  If GrdSL.Rows - 1 > 0 Then GrdSL.SetFocus
  
  Set wRs_ShwLd = Nothing
End Sub
Private Function NextDate(ByVal mi_GapDays As Integer, ByVal ms_StrtDt As String) As String
  Dim wDtFlag As Boolean, ws_Dt As String
  wDtFlag = False
  ws_Dt = DateAdd("d", mi_GapDays, ms_StrtDt)
  Do While wDtFlag = False
    wDtFlag = IIF(moCn.GetFldVal("Select DtHolYN from DtTbl where DtTag= 'HOL' and " + _
                 "DtDt= '" + ws_Dt + "' ") = "Y", False, True)
    If wDtFlag = False Then ws_Dt = DateAdd("d", 1, ws_Dt)
  Loop
  NextDate = ws_Dt
End Function
Private Function DateVal(ByVal ms_Dt As String) As String
  If ms_Dt = "" Then DateVal = 0: Exit Function
  'DateVal = Val(CStr(Year(ms_Dt)) + Format(Month(ms_Dt), "00") + Format(Day(ms_Dt), "00"))
  DateVal = CStr(Year(ms_Dt)) + Format(Month(ms_Dt), "00") + Format(Day(ms_Dt), "00")
End Function
Private Sub DispFra(ByVal pv_DispFra As en_OrdPlnFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_DispFra
  Case Is = Scop
    If FraSC.Visible = True Then
      FraSC.Visible = False
      FraSC.Enabled = False
      Call EnaDisaCmds(False, CmdSC, Scop)
      CmdSC.SetFocus
    Else
      FraSC.Visible = True
      FraSC.Enabled = True
      ADC("wScOpTcFr").SetFocus
      FraSC.ZOrder
      Call EnaDisaCmds(True, CmdSC)
    End If
  Case Is = ResPln
    If FraResPln.Visible = True Then
      FraResPln.Visible = False
      FraResPln.Enabled = False
      Call EnaDisaCmds(False, CmdResPln, ResPln)
      CmdResPln.SetFocus
    Else
      FraResPln.Visible = True
      FraResPln.Enabled = True
      ADC("wRpOpTcFr").SetFocus
      FraResPln.ZOrder
      Call EnaDisaCmds(True, CmdResPln)
    End If
  Case Is = UpdStat
    If FraUpdStat.Visible = True Then
      FraUpdStat.Visible = False
      FraUpdStat.Enabled = False
      Call EnaDisaCmds(False, CmdUpdStat, UpdStat)
      CmdUpdStat.SetFocus
    Else
      FraUpdStat.Visible = True
      FraUpdStat.Enabled = True
      ADC("wUsOpTcFr").SetFocus
      FraUpdStat.ZOrder
      Call EnaDisaCmds(True, CmdUpdStat)
    End If
  Case Is = GenPln
    If FraGenPln.Visible = True Then
      FraGenPln.Visible = False
      FraGenPln.Enabled = False
      Call EnaDisaCmds(False, CmdGenPln, GenPln)
      CmdGenPln.SetFocus
    Else
      FraGenPln.Visible = True
      FraGenPln.Enabled = True
      ADC("wGpOpTcFr").SetFocus
      FraGenPln.ZOrder
      Call EnaDisaCmds(True, CmdGenPln)
    End If
    
  '*** (Jen 2.14)
  Case Is = AutoGenPln
    If FraAutoGenPln.Visible = True Then
      FraAutoGenPln.Visible = False
      FraAutoGenPln.Enabled = False
      Call EnaDisaCmds(False, CmdAutoGenPln, AutoGenPln)
      CmdAutoGenPln.SetFocus
    Else
      FraAutoGenPln.Visible = True
      FraAutoGenPln.Enabled = True
      ADC("wAgOpTcFr").SetFocus
      FraAutoGenPln.ZOrder
      Call EnaDisaCmds(True, CmdAutoGenPln)
    End If
  '*** (Jen 2.14)
  
  Case Is = ShwLd
    If FraShwLd.Visible = True Then
      FraShwLd.Visible = False
      FraShwLd.Enabled = False
      Call EnaDisaCmds(False, CmdShwLd, ShwLd)
      CmdShwLd.SetFocus
    Else
      FraShwLd.Visible = True
      FraShwLd.Enabled = True
      ADC("wSlPlnDtFr").SetFocus
      FraShwLd.ZOrder
      Call EnaDisaCmds(True, CmdShwLd)
    End If
  Case Is = DsgDet
    If FraDsgDet.Visible = True Then
      FraDsgDet.Visible = False
      FraDsgDet.Enabled = False
      '*** required to pass the array position DsgDet-5, to resotre the 'Causes Validation' State
      Call EnaDisaCmds(False, CmdDsgDet, DsgDet)
      GrdOp.SetFocus
    Else
      FraDsgDet.Visible = True
      FraDsgDet.Enabled = True
      FraDsgDet.ZOrder
      Call EnaDisaCmds(True, CmdDsgDet)
      CmdDsgDet.SetFocus
    End If
  Case Is = RedoPln
    If FraRedoPln.Visible = True Then
      FraRedoPln.Visible = False
      FraRedoPln.Enabled = False
      Call EnaDisaCmds(False, CmdRedoPln, RedoPln)
      GrdOpd.SetFocus
      'CmdRedoPln.SetFocus
    Else
      FraRedoPln.Visible = True
      FraRedoPln.Enabled = True
      CmdRedoPlnGo.SetFocus
      FraRedoPln.ZOrder
      Call EnaDisaCmds(True, CmdRedoPln)
      CmdRedoPln.SetFocus
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1, Optional ByVal pv_DispFra As en_OrdPlnFra)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons

  '*** storing the state of all cmd buttons, when the frame is going to be visible.
  Dim X As Integer
  If pv_ShowFra = True Then
    Call StoreState
  End If

  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'        pv_CmdBut.BackColor = vbWhite
'        pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
        If pv_CmdBut.Pic = 0 Then
          pv_CmdBut.BackColor = vbWhite
          'pv_CmdBut.FontBold = True
        End If
        pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'        pv_CmdBut.BackColor = &H8000000F
'        pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
        If pv_CmdBut.Pic = 0 Then
          pv_CmdBut.BackColor = &H8000000F
          'pv_CmdBut.FontBold = False
        End If
        '*** restoring the 'Causes Validation' state of the current pv_cmdbutton
        If Not IsMissing(pv_DispFra) Then
          pv_CmdBut.CausesValidation = mArr_Cv(pv_DispFra)
        End If
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  'ADC.AllowDelete = Not pv_ShowFra
  
  If pv_ShowFra = True Then
    '** All command buttons are disabled
    CmdSC.Enabled = Not pv_ShowFra
    CmdResPln.Enabled = Not pv_ShowFra
    CmdUpdStat.Enabled = Not pv_ShowFra
    CmdGenPln.Enabled = Not pv_ShowFra
    CmdShwLd.Enabled = Not pv_ShowFra
    CmdDsgDet.Enabled = Not pv_ShowFra
    CmdRedoPln.Enabled = Not pv_ShowFra
    CmdAutoGenPln.Enabled = Not pv_ShowFra      '*** (Jen 2.14)
  Else
    '*** Restoring Ena/Disa state of all Command Buttons
    CmdSC.Enabled = mArr_Ed(Scop)
    CmdResPln.Enabled = mArr_Ed(ResPln)
    CmdUpdStat.Enabled = mArr_Ed(UpdStat)
    CmdGenPln.Enabled = mArr_Ed(GenPln)
    CmdShwLd.Enabled = mArr_Ed(ShwLd)
    CmdDsgDet.Enabled = mArr_Ed(DsgDet)
    CmdRedoPln.Enabled = mArr_Ed(RedoPln)
    CmdAutoGenPln.Enabled = mArr_Ed(AutoGenPln)   '*** (Jen 2.14)
  End If

  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraSC.Visible = False
  FraSC.Enabled = False
  
  FraResPln.Visible = False
  FraUpdStat.Visible = False
  FraGenPln.Visible = False
  FraShwLd.Visible = False
  FraDsgDet.Visible = False
  FraRedoPln.Visible = False
  FraAutoGenPln.Visible = False '*** (Jen 2.14)
End Sub
Private Sub StoreState()
  '*** This procedure is used to store the CausesValidation State and the Enabled/Disabled
      'states of all the Command Buttons
  '*** This is used for restoring the original states in the future
  mArr_Cv(Scop) = CmdSC.CausesValidation
  mArr_Cv(ResPln) = CmdResPln.CausesValidation
  mArr_Cv(UpdStat) = CmdUpdStat.CausesValidation
  mArr_Cv(GenPln) = CmdGenPln.CausesValidation
  mArr_Cv(ShwLd) = CmdShwLd.CausesValidation
  mArr_Cv(DsgDet) = CmdDsgDet.CausesValidation
  mArr_Cv(RedoPln) = CmdRedoPln.CausesValidation
  mArr_Cv(AutoGenPln) = CmdAutoGenPln.CausesValidation    '*** (Jen 2.14)
  
  mArr_Ed(Scop) = CmdSC.Enabled
  mArr_Ed(ResPln) = CmdResPln.Enabled
  mArr_Ed(UpdStat) = CmdUpdStat.Enabled
  mArr_Ed(GenPln) = CmdGenPln.Enabled
  mArr_Ed(ShwLd) = CmdShwLd.Enabled
  mArr_Ed(DsgDet) = CmdDsgDet.Enabled
  mArr_Ed(RedoPln) = CmdRedoPln.Enabled
  mArr_Ed(AutoGenPln) = CmdAutoGenPln.Enabled     '*** (Jen 2.14)
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
  ws_BtnStr = "CmdGenPlnGo,CmdRedoPlnGo,CmdResPlnGo,CmdShwLdGo,CmdUpdStatGo,CmdAutoGenPlnGo"
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
Private Sub GrdOp_Validate(Cancel As Boolean)
  Cancel = GrdOp.Validate
End Sub

