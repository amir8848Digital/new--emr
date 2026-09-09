VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvEPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Export Performance Statement"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11730
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5835
      TabIndex        =   39
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   40
         Top             =   135
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOUSRCD"
         IdName          =   "UOUSRCD"
      End
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   1
         Left            =   1860
         TabIndex        =   41
         Top             =   150
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         MaxLength       =   0
         DataField       =   "UOMNUCD"
         IdName          =   "UOMNUCD"
      End
   End
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7050
      TabIndex        =   38
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   35
         Top             =   -165
         Visible         =   0   'False
         Width           =   495
         lastProp        =   500
         _cx             =   5080
         _cy             =   5080
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      End
   End
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   29
      Top             =   9660
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   979
      AllowDelete     =   0   'False
      Automate        =   -1  'True
      EntType         =   1
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame fra_nkeys 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   0
      TabIndex        =   34
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   33
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvEPS.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvEPS.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   37
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   14985
            Begin VB.OptionButton OptInvDt 
               Caption         =   "Print Date"
               BeginProperty Font 
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
               Left            =   4680
               TabIndex        =   28
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   4200
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.OptionButton OptInvDt 
               Caption         =   "Invoice Date"
               BeginProperty Font 
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
               Left            =   1920
               TabIndex        =   27
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   4200
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "VA% on Cost"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Tag             =   "Annex"
               ToolTipText     =   "Check To apply Value Addition percent on Cost"
               Top             =   3840
               Width           =   1980
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   1920
               TabIndex        =   14
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   2570
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2700
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1080
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "InYy >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   1920
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "InTc >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   3120
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "InChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   1920
               TabIndex        =   7
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "InNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   1920
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   795
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   4680
               TabIndex        =   15
               ToolTipText     =   "EnterTo  Raw Material Category"
               Top             =   2570
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "RmCtg <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   13
               Left            =   7290
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   2570
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "RmCtg In "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   5460
               TabIndex        =   5
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1080
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "InYy <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   4680
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "InTc <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   5
               Left            =   5880
               TabIndex        =   6
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1080
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "InChr <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   4680
               TabIndex        =   8
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "InNo <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   4680
               TabIndex        =   10
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   1920
               TabIndex        =   9
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   11
               ToolTipText     =   "Enter Export Number Selection"
               Top             =   1970
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   2640
               TabIndex        =   55
               ToolTipText     =   "Annexure Attached (Y/N)"
               Top             =   3820
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   1920
               TabIndex        =   17
               ToolTipText     =   "Enter From Import Type Scope"
               Top             =   2860
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "TImpTyp >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   28
               Left            =   4680
               TabIndex        =   18
               ToolTipText     =   "Enter To Import Type Scope"
               Top             =   2860
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "TImpTyp <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   29
               Left            =   7290
               TabIndex        =   19
               ToolTipText     =   "Enter Import Type Selection"
               Top             =   2860
               Visible         =   0   'False
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "TImpTyp In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   44
               Left            =   4680
               TabIndex        =   21
               ToolTipText     =   "Enter To Supplier"
               Top             =   3160
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "TCmCd <="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   7290
               TabIndex        =   22
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   3160
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "TCmCd In "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   52
               Left            =   1920
               TabIndex        =   20
               ToolTipText     =   "Enter From Supplier"
               Top             =   3160
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "TCmCd >= "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   4680
               TabIndex        =   24
               ToolTipText     =   "Enter To Supplier"
               Top             =   3470
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOVACTGTO"
               IdName          =   "UOVACTGTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   7290
               TabIndex        =   25
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   3470
               Visible         =   0   'False
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOVACTGSEL"
               IdName          =   "UOVACTGSEL"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   12
               Left            =   1920
               TabIndex        =   23
               ToolTipText     =   "Enter From Supplier"
               Top             =   3470
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOVACTGFR"
               IdName          =   "UOVACTGFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   80
               Left            =   7185
               TabIndex        =   60
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   4200
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   14
               Left            =   4680
               TabIndex        =   13
               ToolTipText     =   "Enter From Purchase Voucher Date"
               Top             =   2280
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "TDt<="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   15
               Left            =   1920
               TabIndex        =   12
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   2280
               Visible         =   0   'False
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "TDt >="
            End
            Begin VB.Label LblConsDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Consider Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Top             =   4200
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label LblPurDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Pur Vch Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   59
               Top             =   2280
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblVaCtg 
               BackStyle       =   0  'Transparent
               Caption         =   "VA Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   58
               Top             =   3520
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.Label LblSuppCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   57
               Top             =   3200
               Width           =   1425
            End
            Begin VB.Label LblImpTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Import Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   56
               Top             =   2900
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Date/Print Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   19
               Left            =   120
               TabIndex        =   54
               Top             =   1665
               Width           =   1785
            End
            Begin VB.Label LblSlashChrTo 
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5820
               TabIndex        =   53
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label LblSlashYyTo 
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5400
               TabIndex        =   52
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label LblInvNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   51
               Top             =   1380
               Width           =   1815
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   6
               Left            =   1890
               TabIndex        =   50
               Top             =   330
               Width           =   1335
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   2
               Left            =   4680
               TabIndex        =   49
               Top             =   330
               Width           =   1275
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Selection"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   3
               Left            =   7320
               TabIndex        =   48
               Top             =   330
               Width           =   1575
            End
            Begin VB.Label LblExpNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp No(s)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   47
               Top             =   1970
               Width           =   1065
            End
            Begin VB.Label ALBL 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   15
               Left            =   120
               TabIndex        =   46
               Top             =   795
               Width           =   1335
            End
            Begin VB.Label LblSlashYyFr 
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2640
               TabIndex        =   45
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label LblSlashChrFr 
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3060
               TabIndex        =   44
               Top             =   1080
               Width           =   105
            End
            Begin VB.Label LblInvTcYy 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   43
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label LblRmCtg 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   42
               ToolTipText     =   "Location"
               Top             =   2590
               Width           =   1425
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvEPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepInvEPS
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
   ' Sachin 3.11.0.1
   ' Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, space(45) as DGrp1, space(45) as DGrp2,
   ' space(45) as DGrp3, space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc, space(30) as qRefDocs,
   ' TzPurAmt as qMetImp, InCoCd, InSelfName, Cast(0 as int) qTrans, InSbNo, TzRmZCd, InSbDt, TzRmQty as IdQty,
   ' TzRmWt, TzRmWt as qTotWt,  TzPurAmt, TzPurAmt as qStudCst, TzPurAmt as qXCst, TzPurAmt as qTotCst, InFOB,
   ' InFOB As qValAdd, InFOB As qPct, InCmFinDest
   ' From InvHd
   ' Join TxndZ On TzCoCd=InCoCd and TzTc=InTc and TzYy=InYY and TzChr=InChr and TzNo=InNo
   ' Where 1 = 2
' *** Report Sql ***
  
  Dim wrepcnd As String, wCnd As String
  Dim wSqlStrA As String, wSqlStrB As String, wSqlStrC As String, wSqlStrC2 As String
  Dim grpflds As String
  Dim wExpNo As String, wInvCurCndStr As String
  
  Dim wVAOnCstStr As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
 
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  '3.11.2
  Dim wInvDt As String
  wInvDt = "InDt "
  If UCase(adc.MenuCd) <> UCase("RepEPSPlainJwl") Then
    If adc("UoYN2") = "Y" Then
      wInvDt = "InDt "
    Else
      wInvDt = "InPrnDt "
    End If
      adc("UoDtFr").CmpStr = wInvDt + ">= "
      adc("UoDtTo").CmpStr = wInvDt + "<= "
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd
    
  Dim wGrp(6) As String, i As Integer, wGrndPos As Integer, wDetPos As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0: wGrndPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Grand)" Then
      wGrndPos = i: Exit For
    End If
    If wGrp(i) = "(None)" Then
      wNonePos = i: Exit For
    End If
  Next i
  
  '*** If Detail option is not selected then report is not shown ***
  'If wDetPos = 0 Then DispMsg "Detail Option Has To Be Selected For This Report", etError  '3.11.2
  
  If wGrndPos <> 0 Then      '3.11.2
    Rep.ProfA.Suppress = True
    Rep.ProfB.Suppress = True
  End If
        
  '3.11.2 - All Invoices should be of the same Currency Code ***
  Dim wChkRes As MDORowSet
  wInvCurCndStr = IIF(adc("UoInTcFr") <> "", " And InTc >= '" + adc("UoInTcFr") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInTcTo") <> "", " And InTc <= '" + adc("UoInTcTo") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInYyFr") <> "", " And InYy >= '" + adc("UoInYyFr") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInYyTo") <> "", " And InYy <= '" + adc("UoInYyTo") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInChrFr") <> "", " And InChr >= '" + adc("UoInChrFr") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInChrTo") <> "", " And Inchr <= '" + adc("UoInChrTo") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInNoFr") <> "" And adc("UoInNoFr") <> "0", " And InNo >= " + CStr(adc("UoInNoFr")) + " ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInNoTo") <> "" And adc("UoInNoFr") <> "0", " And InNo <= " + CStr(adc("UoInNoTo")) + " ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoDtFr") <> "01/01/80" And adc("UoDtFr") <> "", " And " + wInvDt + " >= '" + CStr(adc("UoDtFr")) + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoDtTo") <> "01/01/80" And adc("UoDtTo") <> "", " And " + wInvDt + " <= '" + CStr(adc("UoDtTo")) + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoInExpNoSel") <> "", " And InExpNo In (" + adc("UoInExpNoSel") + ") ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoVaCtgFr") <> "", " And IdVaCtg >= '" + adc("UoVaCtgFr") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoVaCtgTo") <> "", " And IdVaCtg <= '" + adc("UoVaCtgTo") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoVaCtgSel") <> "", " And IdVaCtg In (" + adc("UoVaCtgSel") + ") ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoCmCdFr") <> "", " And InCmCd >= '" + adc("UoCmCdFr") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoCmCdTo") <> "", " And InCmCd <= '" + adc("UoCmCdTo") + "' ", "")
  wInvCurCndStr = wInvCurCndStr + IIF(adc("UoCmCdSel") <> "", " And InCmCd In (" + adc("UoCmCdSel") + ") ", "")
   
  '6.1
  Set wChkRes = moCn.OpenRes(" Set Dateformat " + DtFmtStr() + " Select 'x' from InvDsg " + _
                  "     Join InvHd On InIdNo=IdInIdNo And InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo And InPrtKey=IdPrtKey " + _
                  "     Join CustMst On CmCtg='C' And CmCd=InCmCd " + _
                  " Where 1=1 " + wInvCurCndStr + _
                  " Having Min(CmCurCd)<>Max(CmCurCd)")
  If wChkRes.RecCount > 0 Then
        DispMsg "All Invoices in the given scope should be of the Same Currency Code.", etError
        Exit Sub
  End If
  
  Rep.RF1.Suppress = True
  Rep.RF2.Suppress = True
  Rep.PaperOrientation = crPortrait
  Rep.RHPortrait.Suppress = False
  
  Select Case UCase(adc.MenuCd)
  ' PROFORMA A
  Case UCase("RepEPSPlainJwl")
  Dim wsql As String, wsql1 As String
    adc("UoCoCdFr").CmpStr = "InCoCd = ": adc("UoDtFr").CmpStr = wInvDt + " >=": adc("UoDtTo").CmpStr = wInvDt + " <="
    adc("UoRmCtgSel").CmpStr = "": adc("UoRmCtgFr").CmpStr = "": adc("UoRmCtgTo").CmpStr = ""
    adc("UoPmCdFr").CmpStr = "TImpTyp >=": adc("UoPmCdTo").CmpStr = "TImpTyp <=": adc("UoPmCdSel").CmpStr = "TImpTyp In "
    adc("UoInTcFr").CmpStr = "InTc >=": adc("UoInTcTo").CmpStr = "InTc <=": adc("UoInYyFr").CmpStr = "InYy >=": adc("UoInYyTo").CmpStr = "InYy <="
    adc("UoInChrFr").CmpStr = "InChr >=": adc("UoInChrTo").CmpStr = "InChr <=": adc("UoInNoFr").CmpStr = "InNo >=": adc("UoInNoTo").CmpStr = "InNo <="
    adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = "": adc("UoInExpNoSel").CmpStr = "InExpNo In "

    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd

    '3.12.0
    Dim wVaCtgCnd, wVaCtgCndStr As String
    wVaCtgCnd = IIF(adc("UoVaCtgFr") <> "", " and IdVaCtg>='" + adc("UoVaCtgFr") + "'", "")
    wVaCtgCnd = wVaCtgCnd + IIF(adc("UoVaCtgTo") <> "", " and IdVaCtg<='" + adc("UoVaCtgTo") + "'", "")
    wVaCtgCnd = wVaCtgCnd + IIF(adc("UoVaCtgSel") <> "", " and IdVaCtg In (" + adc("UoVaCtgSel") + ")", "")
    wVaCtgCndStr = IIF(wVaCtgCnd <> "", " and IrRmZCd In " + _
                "         (Select Distinct Ir.IrRmZCd " + _
                "          From InvRm Ir " + _
                "             Join InvDsg Id On Id.IdIdNo=Ir.IrIdIdNo and Id.IdCoCd= Ir.IrCoCd and Id.IdTc= Ir.IrTc and Id.IdYy= Ir.IrYy and Id.IdChr= Ir.IrChr and Id.IdNo= Ir.IrNo and Id.IdSr= Ir.IrSr  and Id.IdPrtKey=Ir.IrPrtKey " + _
                "          Where InIdNo = Id.IdInIdNo And InCoCd = Id.IdCoCd And InTc = Id.IdTc And InYy = Id.IdYy And InChr = Id.IdChr And InNo = Id.IdNo And InPrtKey = Id.IdPrtKey " + _
                 wVaCtgCnd + ") ", "")
                 
                 
    Dim wMetWtStr As String, wLsWtStr As String
    wMetWtStr = "IrRmZWt*IsNull(x1.RmPurityZ,Rm.RmPurityZ)/(case when z.RmPurityZ> 0 then z.RmPurityZ else 1 end)"
    wLsWtStr = "(IdCustmMtchWt- IrRmZWt)*Rm.RmPurityZ/(case when z.RmPurityZ> 0 then z.RmPurityZ else 1 end)"
                 '6.1
    wsql = " Set Dateformat " + DtFmtStr() + " Select IDENTITY(INT, 1, 1) AS qIrSrNo, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, max(IdQty) as IdQty, max(IdGldLs) as IdGldLs,  " + _
           "    sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) / (case when Rm.RmCtg in ('D', 'C') then 5 else 1 end) " + _
           "    * (case when Rm.RmCtg in ('G', 'P', 'S', 'L') then Rm.RmPurityZ " + _
           "    * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100)/ (case when Z.RmPurityZ> 0 then Z.RmPurityZ else 1 end) else 1 end)) as qIrWt, " + _
           "    sum(IrRmZWt/(case when Rm.RmCtg in ('D', 'C') then 5 else 1 end)) as qGrsWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then 1 else 0 end)) as qNetWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldNetWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlNetWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlNetWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdNetWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('M') then 1 else 0 end)) as qMNetWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then 1 else 0 end)) as qNetPureWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdWt, " + _
           "    sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('M') then 1 else 0 end)) as qMWt, " + _
           "    sum((case when IrMainMet= 'Y' then (IdCustmMtchWt - IrRmZWt) else (IrRmZWt* IdGldLs/100) end)* (case when Rm.RmCtg in ('G','P','S','L') then 1 else 0 end)) as qLsWt, " + _
           "    sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('G','P','S','L') then 1 else 0 end)) as qPureLsWt, " + _
           "    sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldLsWt, " + _
           "    sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlLsWt, " + _
           "    sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlLsWt, " + _
           "    sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdLsWt, "

    wsql1 = wsql + " sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('X') then 1 else 0 end)) as qXWt, " + _
           "    sum(IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as qDCSWt, " + _
           "    sum(IrRmQty*(case when Rm.RmCtg in ('X') then 1 else 0 end)) as qXIrRmQty, " + _
           " max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, IrRmZCd, max(Z.RmDesc) as qTyp, sum(IrRmQty*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as IrRmQty, " + _
           " IrRmZCd+ Space(12-Len(IrRmZCd))+ Str(IrRmZPureRt, 9, 2) as qIrZCdRt, IrRmZPureRt as qSalRt, " + _
           " sum(IrRmZRt*IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as qDCSSalVal, " + _
           " max(case when IdQty> 0 then IdZFob/IdQty else 0 end) as qAvgSalPrc, max(IdZFob) as qFOB, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, " + _
           " max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, space(2) as OrRmCtg, max(IdTrayNo) as IdTrayNo " + _
           "  into #TMPB " + _
           " From InvRm " + _
           "     join RmMst Rm on  Rm.RmPrtKey=IrPrtKey and  Rm.RmCd= IrRmCd " + _
           "     join RmMst Z on  Z.RmPrtKey=IrPrtKey and  Z.RmCd= (Case when Rm.RmCtg='X' Then (Select RmZCd From RmMst x where x.RmCd=Rm.RmBaseCd) Else  IrRmZCd End) " + _
           "     join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr  and IdPrtKey=IrPrtKey " + _
           "     join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo  and InPrtKey=IdPrtKey " + _
           "     join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr  and OdPrtKey=IdPrtKey " + _
           "     join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= ''  and DmPrtKey=OdPrtKey " + _
           "     Left Outer Join RmMst x1 On x1.RmCd=Rm.RmBaseCd " + _
           "  Where IdTc='IN' And IrPrtKey='C' and  Round(IrRmZWt, 3) > 0  " + wCnd + _
                    wVaCtgCnd + _
           "  Group By IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IrRmZCd, IrRmZPureRt " + _
           "  Order By IdCoCd, IdTc, IdYy, IdChr, IdNo,IdSr, IrRmZCd, qSalRt"
    moCn.CreateTmpTable "#TMPB", wsql1

    adc("UoCoCdFr").CmpStr = "M.TzCoCd=": adc("UoDtFr").CmpStr = "": adc("UoDtTo").CmpStr = ""
    adc("UoRmCtgSel").CmpStr = "": adc("UoRmCtgFr").CmpStr = "": adc("UoRmCtgTo").CmpStr = ""
    adc("UoPmCdFr").CmpStr = "": adc("UoPmCdTo").CmpStr = "": adc("UoPmCdSel").CmpStr = ""
    adc("UoInTcFr").CmpStr = "M.TzTc>=": adc("UoInTcTo").CmpStr = "M.TzTc<=": adc("UoInYyFr").CmpStr = "M.TzYy>=": adc("UoInYyTo").CmpStr = "M.TzYy<="
    adc("UoInChrFr").CmpStr = "M.TzChr>=": adc("UoInChrTo").CmpStr = "M.TzChr<=": adc("UoInNoFr").CmpStr = "M.TzNo>=": adc("UoInNoTo").CmpStr = "M.TzNo<="
    adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = "": adc("UoInExpNoSel").CmpStr = ""

    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd
    
    wVaCtgCndStr = IIF(wVaCtgCnd <> "", " and M.TzRmZCd In " + _
                "           ( Select Distinct Ir.IrRmZCd From InvRm Ir " + _
                "               Join InvDsg Id On Id.IdIdNo=Ir.IrIdIdNo and Id.IdCoCd= Ir.IrCoCd and Id.IdTc= Ir.IrTc and Id.IdYy= Ir.IrYy and Id.IdChr= Ir.IrChr and Id.IdNo= Ir.IrNo and Id.IdSr= Ir.IrSr  and Id.IdPrtKey=Ir.IrPrtKey " + _
                "            Where M.TzCoCd = Id.IdCoCd And M.TzTc = Id.IdTc And M.TzYy = Id.IdYy And M.TzChr = Id.IdChr And M.TzNo = Id.IdNo And M.TzPrtKey = Id.IdPrtKey " + _
                wVaCtgCnd + ") ", "")
'6.1
    wsql = "Set Dateformat " + DtFmtStr() + " Select M.TzCoCd, M.TzTc, M.TzYy, M.TzChr, M.TzNo, IDENTITY(INT, 1, 1) as qTzSrNo, max(TBOENo) as TBOENo, " + _
            "    max(TBOEDt) as TBOEDt, M.TzRmZCd as TzRmZCd, M.TzPurRt as qCstRt, M.TzRmZCd+ Space(12-Len(M.TzRmZCd))+ Str(IzRmZRt, 9, 2) as qTzZCdRt, IzRmZRt, " + _
            "    sum(M.TzRmWt/(case when RmCtg in ('D', 'C') then 5 else 1 end)) as qTzWt " + _
            " Into #TMPA " + _
            " From TxndZ M " + _
            "   join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey= M.TzRefKey and B.TzBM= 'B'  and B.TzPrtKey=M.TzPrtKey " + _
            "   join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy and TChr= B.TzChr and TNo= B.TzNo  and TPrtKey=B.TzPrtKey " + _
            "   join InvZRm on IzCoCd= M.TzCoCd and IzTc= M.TzTc and IzYy= M.TzYy and IzChr= M.TzChr and IzNo= M.TzNo and IzSr= M.TzSr  and IzPrtKey=M.TzPrtKey " + _
            "   join RmMst on RmCd= M.TzRmZCd  and RmPrtKey=M.TzPrtKey " + _
            " Where M.TzTc='IN' " + wCnd + _
            wVaCtgCndStr + _
            " Group By M.TzCoCd, M.TzTc, M.TzYy, M.TzChr, M.TzNo, TBOENo, M.TzRmZCd, IzRmZRt, M.TzPurRt "
    moCn.CreateTmpTable "#TMPA", wsql

    wsql = " Select qIrSrNo, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt,  qGldWt, qPlWt, qSlWt, qPdWt, qMWt, qLsWt, qPureLsWt, " + _
           "    qNetPureWt, qGldNetWt, qPlNetWt, qSlNetWt, qPdNetWt, qMNetWt, qGldLsWt, qPlLsWt, qSlLsWt, qPdLsWt, qXWt, qDCSWt, qXIrRmQty, RmCtg, RmSCtg, " + _
           "    IrRmZCd, qTyp, IrRmQty, qIrZCdRt, qSalRt, qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, " + _
           "    qTzSrNo , TBOENo, TBOEDt, qCstRt, qTzWt " + _
           " into #TMPC " + _
           " From #TMPB " + _
           "    Join #TMPA on IdCoCd=TzCoCd And IdTc=TzTc And IdYy=TzYy And IdChr=TzChr And IdNo=TzNo And qTzZCdRt= qIrZCdRt " + _
           " Order By IdCoCd, IdTc, IdYy, IdChr, IdNo, qIrSrNo, qTzWt, qTzSrNo "
    moCn.CreateTmpTable "#TMPC", wsql

    Call ApportnTbl

    wsql = " Insert into #TMPD " + _
          "  Select qIrSrNo, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, qGldWt, qPlWt, qSlWt, qPdWt, qMWt, qLsWt, qPureLsWt, qNetPureWt, qGldNetWt, qPlNetWt, qSlNetWt, " + _
          "       qPdNetWt, qMNetWt, qGldLsWt, qPlLsWt, qSlLsWt, qPdLsWt, qXWt, qDCSWt, qXIrRmQty, RmCtg, RmSCtg, IrRmZCd, qTyp, IrRmQty, qSalRt, qIrZCdRt, qDCSSalVal, " + _
          "       qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, TBOENo, TBOEDt, qCstRt, " + _
          "       qIrWt- (case when IsNull((Select sum(qTzWt) from #TMPA a " + _
          "                      Where a.TzCoCd = c.IdCoCd And a.TzTc = c.IdTc And a.TzYy = c.IdYy And a.TzChr = c.IdChr And a.TzNo = c.IdNo " + _
          "                              And a.qTzSrNo< c.qTzSrNo and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "            IsNull((Select sum(qIrWt) from #TMPB b " + _
          "                      Where b.IdCoCd = c.IdCoCd And b.IdTc = c.IdTc And b.IdYy = c.IdYy And b.IdChr = c.IdChr And b.IdNo = c.IdNo " + _
          "                              And b.qIrSrNo< c.qIrSrNo and b.qIrZCdRt= c.qIrZCdRt), 0) > 0 then " + _
          "            IsNull((Select sum(qTzWt) from #TMPA a " + _
          "                      Where a.TzCoCd = c.IdCoCd And a.TzTc = c.IdTc And a.TzYy = c.IdYy And a.TzChr = c.IdChr And a.TzNo = c.IdNo " + _
          "                              And a.qTzSrNo< c.qTzSrNo and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "            IsNull((Select sum(qIrWt) from #TMPB b " + _
          "                      Where b.IdCoCd = c.IdCoCd And b.IdTc = c.IdTc And b.IdYy = c.IdYy And b.IdChr = c.IdChr And b.IdNo = c.IdNo " + _
          "                              And b.qIrSrNo< c.qIrSrNo and b.qIrZCdRt= c.qIrZCdRt), 0) else 0 end) as qFinalReq, " + _
          "       qTzWt- (case when IsNull((Select sum(qIrWt) from #TMPB b " + _
          "                      Where b.IdCoCd = c.IdCoCd And b.IdTc = c.IdTc And b.IdYy = c.IdYy And b.IdChr = c.IdChr And b.IdNo = c.IdNo " + _
          "                              And b.qIrSrNo< c.qIrSrNo and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
          "            IsNull((Select sum(qTzWt) from #TMPA a " + _
          "                      Where a.TzCoCd = c.IdCoCd And a.TzTc = c.IdTc And a.TzYy = c.IdYy And a.TzChr = c.IdChr And a.TzNo = c.IdNo " + _
          "                               And a.qTzSrNo< c.qTzSrNo and a.qTzZCdRt= c.qIrZCdRt), 0) > 0 then " + _
          "            IsNull((Select sum(qIrWt) from #TMPB b " + _
          "                      Where b.IdCoCd = c.IdCoCd And b.IdTc = c.IdTc And b.IdYy = c.IdYy And b.IdChr = c.IdChr And b.IdNo = c.IdNo " + _
          "                              And b.qIrSrNo< c.qIrSrNo and b.qIrZCdRt= c.qIrZCdRt), 0) - "

     wsql1 = wsql + "  IsNull((Select sum(qTzWt) from #TMPA a " + _
                    "            Where a.TzCoCd = c.IdCoCd And a.TzTc = c.IdTc And a.TzYy = c.IdYy And a.TzChr = c.IdChr And a.TzNo = c.IdNo " + _
                    "                    And a.qTzSrNo< c.qTzSrNo and a.qTzZCdRt= c.qIrZCdRt), 0) else 0 end) as qFinalAvl " + _
                    "   from #TMPC c "
     moCn.Execute (wsql1)
    
    If UCase(adc("UoYn1")) = "Y" Then
      wVAOnCstStr = "(Case when Sum(qMetCstVal)=0 Then 100 Else Sum(qMetCstVal) End)"
    Else
      wVAOnCstStr = "(Case when Sum(InFOB)=0 Then 100 Else Sum(InFOB) End)"
    End If
'6.1
    wsql = " Set Dateformat " + DtFmtStr() + " Select " + grpflds + ", qInvNo, Max(InSbNo) as InSbNo, Max(InSbDt) as InSbDt, Max(qVaCtg) as qVaCtg, Sum(qQty) as IdQty, " + _
           "     Cast(Sum(qNetWt)+Sum(qXNetWt)+(Sum(qDCSWt)/5) as Decimal(16,3)) as qTotWt, " + _
           "     Cast(Sum(qNetPureWt)+Sum(qXWt)+(Sum(qDCSWt)/5) as Decimal(16,3)) as TzRmWt, " + _
           "     Cast(Sum(qDCCstVal) as Decimal(16,3)) qDCCstVal, " + _
           "     Cast(Sum(qMetCstVal) as Decimal(16,3)) as qMetCst, Cast(Sum(InFob) as Decimal(16,3)) as InFob,  Cast(Sum(qValAddn) as Decimal(16,3)) as qValAdd, " + _
           "     Cast((100*Sum(qValAddn)/" + wVAOnCstStr + ") as Decimal(16,2)) as qPct, Max(InCmFinDest) as InCmFinDest " + _
           " From " + _
           " (Select Max(qInvNo) as qInvNo, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, Max(InSbNo) as InSbNo, Max(InSbDt) as InSbDt,  Max(qVaCtg) as qVaCtg, " + _
           "    Max(qQty) as qQty, Sum(qNetWt) as qNetWt, Sum(qNetPureWt) as qNetPureWt, Sum(qXNetWt) as qXNetWt, Sum(qXWt) as qXWt, Sum(qDCSWt) as qDCSWt, " + _
           "    max(InFob) as InFob, Sum(qDCCstVal) as qDCCstVal, Sum(qMetCstVal) as qMetCstVal, (max(InFob)-Sum(qValAddn)) as qValAddn, Max(InCmFinDest) as InCmFinDest " + _
           " From " + _
           " (Select  (IdTc + '/' +  IdYy +'/'+ IdChr + '/'+ Cast(IdNo as varChar(10))) as qInvNo, IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, Max(InSbNo) as InSbNo, Max(InPrnDt) as InSbDt,  Max(IdVaCtg) as qVaCtg, " + _
           "    Max(IdQty) as qQty, Max(Case When D.RmCtg In ('G','P','S','L','M') Then (qNetWt+qLsWt) Else 0 End) as qNetWt, " + _
           "    Max(Case When D.RmCtg In ('G','P','S','L','M') Then (qNetPureWt+qPureLsWt)  Else 0 End) as qNetPureWt, max(qFob) as InFob, " + _
           "    Sum((qCstRt*((case when qFinalReq>0 and qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)* " + _
           "             (Case When D.RmCtg In ('D','C') Then 5 Else 0 End)))) as qDCCstVal, " + _
           "    Sum(((case when qFinalReq>0 and qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)* " + _
           "        (Case When D.RmCtg In ('X') Then 1 Else 0 End))) as qXNetWt, Max(qXWt) as qXWt, Max(qDCSWt) as qDCSWt, " + _
           "    Sum((qCstRt*((case when qFinalReq>0 and qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)*(Case When D.RmCtg In ('D','C') Then 5 Else 1 End)))) as qMetCstVal, " + _
           "   (Sum((qCstRt*((case when qFinalReq>0 and qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)*(Case When D.RmCtg In ('D','C') Then 5 Else 0 End)))) + " + _
           "    Sum((qCstRt*((case when qFinalReq>0 and qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)*(Case When D.RmCtg In ('D','C','X') Then 0 Else 1 End))))) as qValAddn, Max(InCmFinDest) as InCmFinDest "
           
    wsql = wsql + " from #TmpD D " + _
           "    Join InvHd On InCoCd=IdCoCd And InTc=IdTc And InYy=IdYy And InChr=IdChr And InNo=IdNo " + _
           " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, qIrZCdRt) qVa " + _
           " Group by IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr) qVa1 " + _
           " Group by qInvNo, IdTc, IdYy, IdChr, IdNo Order by IdTc, IdYy, IdChr, IdNo "
    adc.RepRecSource = wsql + ctMaxDopOpt
                    
    'QUERY BEFORE 3.11.2
    '        wSqlStrA = " Set Dateformat DMY Select " + grpflds + ", InCoCd, Max(InSelfName) as InSelfName, Max(InSbNo) as InSbNo,TzRmZCd,Max(InPrnDt) as InSbDt, " + _
    '                 "   (Select Sum(IdQty) From InvDsg where IdInIdNo=Max(InIdNo) and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and " + _
    '                 "   IdChr= InChr and IdNo= InNo and IdPrtKey=Max(InPrtKey))  as IdQty, Sum(TzRmWt) as TzRmWt, " + _
    '                 "   (Select Sum(IrRmZWt) from InvRm Where IrCoCd=InCoCd And IrTc=InTc And IrYy=InYy And IrChr=InChr And IrNo=InNo And IrRmZCd=TzRmZCd) qTotWt, " + _
    '                 "   Sum(TzPurAmt) as qMetCst, Max(InFOB) as InFOB,Max(InFOB)-Sum(TzPurAmt) as qValAdd, " + _
    '                 "   ((Max(InFOB)-Sum(TzPurAmt))/" + wVAOnCstStr + "*100) as qPct, Max(InCmFinDest) as InCmFinDest " + _
    '                 "   From InvHd " + _
    '                 "      Join TxndZ On TzCoCd=InCoCd and TzTc=InTc and TzYy=InYY and TzChr=InChr and TzNo=InNo " + _
    '                 "      Join RmMst On TzRmZCd=RmCd and RmZ='Y' " + _
    '                 "   Where InTc='IN' " + wCnd + _
    '                 "   Group By InCoCd, InTc, InYy, InChr, InNo, TzRmZCd"
    '   adc.RepRecSource = wSqlStrA + ctMaxDopOpt

    If UCase(adc("UoYn1")) = "Y" Then
        GRP_REP.SetFormula Rep, "wValAddPctA", "100*Sum ({rdo.qValAdd})/(if (Sum ({rdo.qMetCst})>0) then Sum ({rdo.qMetCst}) Else 1)"
    Else
        GRP_REP.SetFormula Rep, "wValAddPctA", "100*Sum ({rdo.qValAdd})/(if (Sum ({rdo.InFOB})>0) then Sum ({rdo.InFOB}) Else 1)"
    End If
    
    Rep.PaperOrientation = crLandscape
    Rep.PHProformaB.Suppress = True: Rep.PHProformaC.Suppress = True
    Rep.ProfB.Suppress = True: Rep.ProfC.Suppress = True
    Rep.RHMain.Suppress = False: Rep.RF1.Suppress = False
    Rep.RHPortrait.Suppress = True
    
  ' PROFORMA B
  Case UCase("RepEPSStudJwl")
      If UCase(adc("UoYn1")) = "Y" Then
        wVAOnCstStr = "((Max(InFOB)-Sum(TzPurAmt))/(Case when Sum(TzPurAmt)=0 Then 1 Else Sum(TzPurAmt) End)*100)"
      Else
        wVAOnCstStr = "((Max(InFOB)-Sum(TzPurAmt))/(Case when Max(InFOB)=0 Then 1 Else Max(InFOB) End)*100)"
      End If
        
      adc("UoCmCdFr").CmpStr = "InCmCd>=": adc("UoCmCdTo").CmpStr = "InCmCd<=": adc("UoCmCdSel").CmpStr = "InCmCd In "
      wrepcnd = adc.RepCond
       wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd
      '6.1
    wSqlStrB = " Set Dateformat " + DtFmtStr() + " Select " + grpflds + ",InTc+'\'+InYy+'\'+InChr+'\'+Cast(InNo as varChar(10)) as qInvNoB, " + _
               "     Sum((Case When RmCtg in ('G','P','S','L','M') Then TzPurAmt Else 0 End )) as qMetCst, " + _
               "     Sum((Case When RmCtg in ('D','C') Then TzPurAmt Else 0 End )) as qDCCst, " + _
               "     Sum((Case When RmCtg ='X' Then TzPurAmt Else 0 End )) as qXCst, " + _
               "     Sum(TzPurAmt) as qTotCst, Max(InFOB) as InFOB,Max(InFOB)-Sum(TzPurAmt) as qValAdd, " + _
               "    " + wVAOnCstStr + " qPct " + _
               "  From InvHd " + _
               "    Join TxndZ On TzCoCd=InCoCd and TzTc=InTc and TzYy=InYY and TzChr=InChr and TzNo=InNo " + _
               "    Join RmMst On TzRmZCd=RmCd and RmZ='Y' " + _
               "  Where TzBM='M' and InTc='IN' " + wCnd + _
               "  Group By InCoCd,InTc,InYy,InChr,InNo Order by InCoCd,InTc,InYy,InChr,InNo"
               
    adc.RepRecSource = wSqlStrB + ctMaxDopOpt
    
    If UCase(adc("UoYn1")) = "Y" Then
        GRP_REP.SetFormula Rep, "wValAddPctA", "100*Sum ({rdo.qValAdd})/Sum ({rdo.qTotCst})"
    Else
        GRP_REP.SetFormula Rep, "wValAddPctA", "100*Sum ({rdo.qValAdd})/Sum ({rdo.InFOB})"
    End If
    
    Rep.PHProformaA.Suppress = True:  Rep.PHProformaC.Suppress = True: Rep.ProfA.Suppress = True: Rep.ProfC.Suppress = True
    Rep.RF2.Suppress = False: Rep.RHMain.Suppress = True
    Rep.txtPgBox.Left = 4970: Rep.txtPgBox.Width = 1352: Rep.TxtMwRepFt.Left = 6700
    Rep.TxtPg.Left = 5051: Rep.TxtPgNo.Left = 5547: Rep.TxtSlash.Left = 5847: Rep.TxtTotPg.Left = 5971
    Rep.TxtProforma.SetText ("PROFORMA 'B'")
    Rep.TxtRepName.SetText ("Statement of Export Performance & Value Addition for Studded/Plain Jewellery")
  
  ' PROFORMA C
  Case UCase("RepEPSMetImpEx")
    adc("UoCoCdFr").CmpStr = "TCoCd=": adc("UoDt1Fr").CmpStr = "TDt>=": adc("UoDt1To").CmpStr = "TDt<="
    adc("UoDtFr").CmpStr = "": adc("UoDtTo").CmpStr = ""
    adc("UoRmCtgSel").CmpStr = "TzRmCtg In ": adc("UoRmCtgFr").CmpStr = "TzRmCtg>=": adc("UoRmCtgTo").CmpStr = "TzRmCtg<="
    adc("UoPmCdFr").CmpStr = "TImpTyp>=": adc("UoPmCdTo").CmpStr = "TImpTyp<=": adc("UoPmCdSel").CmpStr = "TImpTyp In "
    adc("UoInTcFr").CmpStr = "": adc("UoInTcTo").CmpStr = "": adc("UoInYyFr").CmpStr = "": adc("UoInYyTo").CmpStr = ""
    adc("UoInChrFr").CmpStr = "": adc("UoInChrTo").CmpStr = "": adc("UoInNoFr").CmpStr = "": adc("UoInNoTo").CmpStr = ""
    adc("UoCmCdFr").CmpStr = "TSuppCd>=": adc("UoCmCdTo").CmpStr = "TSuppCd<=": adc("UoCmCdSel").CmpStr = "TSuppCd In "
    adc("UoInExpNoSel").CmpStr = ""
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd
    
    Dim wStr As String
    wStr = " Select " + grpflds + ", '' as qRefDocs "
    '6.1
    wSqlStrC = " Set Dateformat " + DtFmtStr() + " " + _
               " Select TTc + '\' + TYy+ '\'+TChr+'\'+Cast(TNo as varChar(10)) as qVchNo, Max(TDt) as TDt,TBOENo as qBOENo,Max(TBOEDt) as qBOEDt, TzRmZCd AS qRmZCd, Sum(TzRmWt) as qTzRmWt " + _
               " From TxndZ " + _
               "     Join Txn On TzCoCd=TCoCd and TzTc=TTc and TzYy=TYy and TzChr=TChr and TzNo=TNo " + _
               " Where TzBM='B' and TzRmCtg in ('G','P','S','L','M') " + wCnd + _
               " Group By TCoCd, TTc, TYy, TChr, TNo, TBOENo, TzRmZCd " + _
               " Order by Max(TDt), TCoCd, TTc, TYy, TChr, TNo"
    
    adc("UoCoCdFr").CmpStr = "IrCoCd=": adc("UoDtFr").CmpStr = wInvDt + ">=": adc("UoDtTo").CmpStr = wInvDt + "<="
    adc("UoDt1Fr").CmpStr = "": adc("UoDt1To").CmpStr = ""
    adc("UoRmCtgSel").CmpStr = "R1.RmCtg In ": adc("UoRmCtgFr").CmpStr = "R1.RmCtg>=": adc("UoRmCtgTo").CmpStr = "R1.RmCtg<="
    adc("UoPmCdFr").CmpStr = "": adc("UoPmCdTo").CmpStr = "": adc("UoPmCdSel").CmpStr = ""
    adc("UoInTcFr").CmpStr = "InTc>=": adc("UoInTcTo").CmpStr = "InTc<=": adc("UoInYyFr").CmpStr = "InYy>=": adc("UoInYyTo").CmpStr = "InYy<="
    adc("UoInChrFr").CmpStr = "InChr>=": adc("UoInChrTo").CmpStr = "InChr<=": adc("UoInNoFr").CmpStr = "InNo>=": adc("UoInNoTo").CmpStr = "InNo<="
    adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
    adc("UoInExpNoSel").CmpStr = " InExpNo In "
    
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " And ", " ") + wrepcnd
    wSqlStrC2 = " Select InTc + '\' + InYy+ '\'+InChr+'\'+Cast(InNo as varChar(10)) as qInvNo, InSbNo as qSbNo,IrRmZCd As qRmZCd, Max(InDt) As qSbDt, " + _
                "   sum(IrRmZWt*R1.RmPurityZ / (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) As qRmZWt, " + _
                "   sum((case when IrMainMet= 'Y' then (IdCustmMtchWt- IrRmZWt) else (IrRmZWt* IdGldLs/100) end)*R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) As qLsWt " + _
                " From InvRm " + _
                "    Join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr = IrChr And IdNo = IrNo And IdSr = IrSr And IdPrtKey = IrPrtKey " + _
                "    Join InvHd on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr = InChr And IdNo = InNo And IdPrtKey = InPrtKey " + _
                "    Join RmMst R1 on R1.RmCd= IrRmCd and R1.RmPrtKey=IrPrtKey " + _
                "    Join RmMst R2 on R2.RmCd= R1.RmZCd and R2.RmPrtKey=R1.RmPrtKey " + _
                " Where IrRmZWt> 0  and IrRmCtg in ('G','P','S','L','M') " + wCnd + _
                " Group By InTc, InYy, Inchr, InNo, InSbNo, IrRmZCd, IdGldLs"

    adc.RepRecSource = wStr + ctMaxDopOpt
    Call moCn.RepRes(Rep.SubRepPurSumm.OpenSubreport, wSqlStrC + ctMaxDopOpt)
    Call moCn.RepRes(Rep.SubRepExpSumm.OpenSubreport, wSqlStrC2 + ctMaxDopOpt)
    Rep.PHProformaA.Suppress = True: Rep.PHProformaB.Suppress = True
    Rep.ProfA.Suppress = True: Rep.ProfB.Suppress = True: Rep.RHMain.Suppress = True
    Rep.txtPgBox.Left = 4970: Rep.txtPgBox.Width = 1352: Rep.TxtMwRepFt.Left = 6700
    Rep.TxtPg.Left = 5051: Rep.TxtPgNo.Left = 5547: Rep.TxtSlash.Left = 5847: Rep.TxtTotPg.Left = 5971
    Rep.TxtProforma.SetText ("PROFORMA 'C'")
    Rep.TxtRepName.SetText ("Monthly Statement of Metal Procurement, Consumption & Export")
  End Select
  
  'Rep.TxtFrInvDt.SetText IIF(adc("UoDtFr") <> "01/01/80", adc("UoDtFr"), "")
  'Rep.TxtToInvDt.SetText IIF(adc("UoDtTo") <> "01/01/80", adc("UoDtTo"), "")
  'Rep.TxtRmCtg.SetText adc("UoRmCtgFr")      ' **** Manali 3.03 - Not Required
    
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub
Private Sub OptInvDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN2") = "Y"
Case Is = 1
  adc("UoYN2") = "N"
End Select
End Sub
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1     'Tag = "Percent on Cost"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

'''' std code not to be changed
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
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection
    Set adc.FirNKeyCtl = adc("UoInTcFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_InTcTyp = "IN"
    gs_PTyp = "ImpTyp"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
        

    
    Call SetGroupSort
    
      ' PROFORMA A (Plain Jewellery)
    If UCase(adc.MenuCd) = UCase("RepEPSPlainJwl") Then
       LblExpNo.Visible = True: LblRmCtg.Visible = False
       LblSuppCd.Visible = False: adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False:
       LblVaCtg.Visible = True: adc("UoVaCtgFr").Visible = True: adc("UoVaCtgTo").Visible = True: adc("UoVaCtgSel").Visible = True      '3.12.0
       adc("UoRmCtgFr").Visible = False: adc("UoRmCtgTo").Visible = False: adc("UoInExpNoSel").Visible = True: adc("UoRmCtgSel").Visible = False
       ALBL(19).Caption = "Inv Date"
    Else
       ALBL(19).Caption = "Inv Date/Print Date"
       LblExpNo.Visible = False: adc("UoInExpNoSel").Visible = False
       OptInvDt(0).Visible = True: OptInvDt(1).Visible = True: LblConsDt.Visible = True
       'PROFORMA B & C (Studded Jewellery and Import Export Statement)
       If UCase(adc.MenuCd) = UCase("RepEPSMetImpEx") Then
          ChkBoxArr(1).Visible = False
          LblRmCtg.Visible = True: adc("UoRmCtgSel").Visible = True
          adc("UoRmCtgFr").Visible = True: adc("UoRmCtgTo").Visible = True
          LblImpTyp.Visible = True: adc("UoPmCdFr").Visible = True: adc("UoPmCdTo").Visible = True: adc("UoPmCdSel").Visible = True
          adc("UoDtFr").Visible = True: adc("UoDtTo").Visible = True: adc("UoDt1Fr").Visible = True: adc("UoDt1To").Visible = True
          LblConsDt.Visible = True: OptInvDt.Item(0).Visible = True: OptInvDt.Item(1).Visible = True: LblPurDt.Visible = True
       Else
          LblRmCtg.Visible = False
          'LblSuppCd.Visible = False: adc("UoCmCdFr").Visible = False: adc("UoCmCdTo").Visible = False: adc("UoCmCdSel").Visible = False:
          LblSuppCd = "Inv Cust"
          adc("UoRmCtgFr").Visible = False: adc("UoRmCtgTo").Visible = False: adc("UoRmCtgSel").Visible = False
       End If
    End If
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  gs_CmCtg = "S"
  If UCase(adc.MenuCd) = UCase("RepEPSStudJwl") Then gs_CmCtg = "C"
  gs_PTyp = "ImpTyp"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_InTcTyp = "IN"
  gs_CmCtg = "S"
  If UCase(adc.MenuCd) = UCase("RepEPSStudJwl") Then gs_CmCtg = "C"
  gs_PTyp = "ImpTyp"
  SetProp Me, IdName, When
End Sub
Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  
  adc("UoCoCdFr") = gs_CoCd
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN2") = "Y" Then OptInvDt(0).Value = True
  If adc("UoYN2") = "N" Then OptInvDt(1).Value = True
  
  If adc("UoYN1") = "" Then adc("UoYn1") = "N"
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  
  If Trim(adc("UODT1FR")) = "" Then adc("UODT1FR") = "01/01/80"
  If Trim(adc("UODT1TO")) = "" Then adc("UODT1TO") = "01/01/80"
  
End Sub
Private Sub DispCoNm()
  'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  Rep.wCoCdPortrait.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  ' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  Rep.TxtRepUsrDtTimePortrait.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If adc.HotKeys(KeyCode, Shift) Then
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
    Cancel = atxt(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub

Private Sub ApportnTbl()
  Dim wStr1 As String, wStr2 As String, wStr3 As String
  '*** Creating Temporary Table #TMPD which is the main Apportioning Sql ******
  wStr1 = "Create table #TMPD " + _
          "(  qIrSrNo     integer         null, " + _
          "   IdCoCd      varchar(3)      null, " + _
          "   IdTc        varchar(3)      null, " + _
          "   IdYy        varchar(3)      null, " + _
          "   IdChr       varchar(3)      null, " + _
          "   IdNo        integer         null, " + _
          "   IdSr        integer         null, " + _
          "   IdDmCd      varchar(15)     null, " + _
          "   IdDmSz      varchar(5)      null, " + _
          "   IdQty       float           null, " + _
          "   IdGldLs     float           null, " + _
          "   qIrWt       float           null, " + _
          "   qGrsWt      float           null, " + _
          "   qNetWt      float           null, " + _
          "   qGldWt      float           null, " + _
          "   qPlWt       float           null, " + _
          "   qSlWt       float           null, " + _
          "   qPdWt       float           null, " + _
          "   qMWt        float           null, " + _
          "   qLsWt       float           null, " + _
          "   qPureLsWt   float           null, "
  
  wStr2 = "   qNetPureWt  float           null, " + _
          "   qGldNetWt   float           null, " + _
          "   qPlNetWt    float           null, " + _
          "   qSlNetWt    float           null, " + _
          "   qPdNetWt    float           null, " + _
          "   qMNetWt     float           null, " + _
          "   qGldLsWt    float           null, " + _
          "   qPlLsWt     float           null, " + _
          "   qSlLsWt     float           null, " + _
          "   qPdLsWt     float           null, " + _
          "   qXWt        float           null, " + _
          "   qDCSWt      float           null, " + _
          "   qXIrRmQty   float           null, " + _
          "   RmCtg       varchar(2)      null, " + _
          "   RmSCtg      varchar(5)      null, " + _
          "   IrRmZCd     varchar(12)     null, " + _
          "   qTyp        varchar(50)     null, " + _
          "   IrRmQty     integer         null, " + _
          "   qSalRt      float           null, " + _
          "   qIrZCdRt    varchar(21)     null, " + _
          "   qDCSSalVal  float           null, "

  wStr3 = "   qAvgSalPrc  float           null, " + _
          "   qFOB        float           null, " + _
          "   IdKt        varchar(5)      null, " + _
          "   IdVaCtg     varchar(5)      null, " + _
          "   DmCtg       varchar(5)      null, " + _
          "   DmSalCtg    varchar(8)      null, " + _
          "   DmUom       varchar(5)      null, " + _
          "   OdDmCol     varchar(5)      null, " + _
          "   OrRmCtg     varchar(2)      null, " + _
          "   IdTrayNo    integer         null, " + _
          "   TBOENo      varchar(20)     null, " + _
          "   TBOEDt      smalldatetime   null, " + _
          "   qCstRt      float           null, " + _
          "   qFinalReq   float           null, " + _
          "   qFinalAvl   float           null " + _
          ")"
          
   moCn.CreateTmpTable "#TMPD", wStr1 + wStr2 + wStr3
End Sub

