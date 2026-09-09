VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpJtTxnLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Listing - Jwlry"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10830
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4785
      TabIndex        =   43
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   44
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
         TabIndex        =   45
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
      Left            =   6000
      TabIndex        =   42
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   40
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
      TabIndex        =   36
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
      Height          =   9675
      Left            =   0
      TabIndex        =   39
      Top             =   60
      Width           =   14895
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   38
         Top             =   30
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpJtTxnLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpJtTxnLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8475
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   14535
            Begin VB.Frame FraShowValWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2040
               TabIndex        =   76
               ToolTipText     =   "Show Value / Weight"
               Top             =   4320
               Width           =   3105
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Value"
                  BeginProperty Font 
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
                  Left            =   70
                  TabIndex        =   29
                  ToolTipText     =   "Show Value / Weight"
                  Top             =   0
                  Width           =   1185
               End
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Weight"
                  BeginProperty Font 
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
                  Left            =   1420
                  TabIndex        =   30
                  ToolTipText     =   "Show Value / Weight"
                  Top             =   0
                  Width           =   1185
               End
            End
            Begin VB.Frame FraOptShowDsg 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2040
               TabIndex        =   74
               ToolTipText     =   "Show Design Yes / No"
               Top             =   4650
               Width           =   2745
               Begin VB.OptionButton OptShowDsg 
                  Caption         =   "No"
                  BeginProperty Font 
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
                  Left            =   1420
                  TabIndex        =   32
                  ToolTipText     =   "Show Design Yes / No"
                  Top             =   0
                  Width           =   705
               End
               Begin VB.OptionButton OptShowDsg 
                  Caption         =   "Yes"
                  BeginProperty Font 
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
                  Left            =   70
                  TabIndex        =   31
                  ToolTipText     =   "Show Design Yes / No"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value   "
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
               Index           =   2
               Left            =   4800
               TabIndex        =   48
               Tag             =   "AddVal3"
               ToolTipText     =   "Check To Add Value"
               Top             =   1680
               Width           =   1410
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Add Value"
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
               Left            =   4800
               TabIndex        =   47
               Tag             =   "AddVal2"
               ToolTipText     =   "Add Value"
               Top             =   1395
               Width           =   1410
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4800
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer"
               Top             =   3385
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7500
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Selection"
               Top             =   3385
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2115
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer"
               Top             =   3385
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2115
               TabIndex        =   13
               ToolTipText     =   "Enter From Order Number"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OmNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4800
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Number"
               Top             =   2250
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OmNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6405
               TabIndex        =   49
               Top             =   1395
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2115
               TabIndex        =   0
               ToolTipText     =   "Enter From Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OmCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3315
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Character"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OmChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2895
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Year"
               Top             =   1965
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OmYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   5985
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Character"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OmChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5565
               TabIndex        =   11
               ToolTipText     =   "Enter To Order Year"
               Top             =   1965
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OmYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   4800
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OmTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   2115
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1965
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OmTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   2115
               TabIndex        =   28
               ToolTipText     =   "Enter Currency Code "
               Top             =   3980
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
               CmpStr          =   "OmCmCurCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   4800
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "OmCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   7500
               TabIndex        =   2
               ToolTipText     =   "Enter Company Code Selection"
               Top             =   540
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OmCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   6405
               TabIndex        =   50
               Top             =   1680
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2115
               TabIndex        =   3
               ToolTipText     =   "Enter Customer Category"
               Top             =   825
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCTGFR"
               IdName          =   "UOCMCTGFR"
               CmpStr          =   "OmCmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2115
               TabIndex        =   4
               ToolTipText     =   "Enter Tc 1"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTC"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2115
               TabIndex        =   5
               ToolTipText     =   "Enter Tc 2"
               Top             =   1395
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC1"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2115
               TabIndex        =   6
               ToolTipText     =   "Enter Tc 3"
               Top             =   1680
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMTC2"
               ReCalcParent    =   "UOOMTCFR"
               IdName          =   "UOOMTC2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4800
               TabIndex        =   16
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "OmDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   2115
               TabIndex        =   15
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   2535
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "OmDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   4800
               TabIndex        =   21
               ToolTipText     =   "Enter To Po Date"
               Top             =   3100
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1TO"
               IdName          =   "UODT1TO"
               CmpStr          =   "OmPoDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2115
               TabIndex        =   20
               ToolTipText     =   "Enter From Po Date"
               Top             =   3100
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODT1FR"
               IdName          =   "UODT1FR"
               CmpStr          =   "OmPoDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4800
               TabIndex        =   18
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2820
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   7500
               TabIndex        =   19
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2820
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   2115
               TabIndex        =   17
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2820
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   4800
               TabIndex        =   26
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   3670
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   7500
               TabIndex        =   27
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   3670
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   2115
               TabIndex        =   25
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   3670
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   5160
               TabIndex        =   72
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   4320
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   5160
               TabIndex        =   75
               ToolTipText     =   "Show Design Yes / No"
               Top             =   4650
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               IdName          =   "WSHOWDSG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2115
               TabIndex        =   33
               ToolTipText     =   "Enter Divided By Amount "
               Top             =   5040
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "UOVALDIVBY"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Divided By"
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
               Index           =   7
               Left            =   90
               TabIndex        =   77
               Top             =   5040
               Width           =   1785
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Design "
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
               Left            =   90
               TabIndex        =   73
               Top             =   4650
               Width           =   1335
            End
            Begin VB.Label LblVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show "
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
               Left            =   90
               TabIndex        =   71
               Top             =   4320
               Width           =   1665
            End
            Begin VB.Label LblSalExec 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   90
               TabIndex        =   70
               Top             =   3670
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO Dt"
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
               Index           =   22
               Left            =   90
               TabIndex        =   69
               Top             =   3100
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Dt"
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
               Left            =   90
               TabIndex        =   68
               Top             =   2520
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "PO No"
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
               Index           =   18
               Left            =   90
               TabIndex        =   67
               Top             =   2820
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 3"
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
               Index           =   10
               Left            =   90
               TabIndex        =   66
               Top             =   1680
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 2"
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
               Left            =   90
               TabIndex        =   65
               Top             =   1395
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tc 1"
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
               Index           =   9
               Left            =   90
               TabIndex        =   64
               Top             =   1110
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Ctg"
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
               Index           =   5
               Left            =   90
               TabIndex        =   63
               Top             =   825
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   175
               Left            =   90
               TabIndex        =   62
               Top             =   3980
               Width           =   1005
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Tc/Yy/Chr"
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
               Index           =   12
               Left            =   90
               TabIndex        =   61
               Top             =   1965
               Width           =   1785
            End
            Begin VB.Label ALBL 
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
               Index           =   0
               Left            =   2835
               TabIndex        =   60
               Top             =   1965
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   4
               Left            =   3255
               TabIndex        =   59
               Top             =   1965
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   3
               Left            =   5505
               TabIndex        =   58
               Top             =   1965
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Index           =   8
               Left            =   5925
               TabIndex        =   57
               Top             =   1965
               Width           =   105
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
               Index           =   21
               Left            =   90
               TabIndex        =   56
               Top             =   540
               Width           =   1335
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
               Index           =   20
               Left            =   7500
               TabIndex        =   55
               Top             =   60
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch No"
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
               Index           =   11
               Left            =   90
               TabIndex        =   54
               Top             =   2250
               Width           =   1485
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   17
               Left            =   90
               TabIndex        =   53
               Top             =   3385
               Width           =   1425
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
               Index           =   1
               Left            =   2115
               TabIndex        =   52
               Top             =   60
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
               Left            =   4800
               TabIndex        =   51
               Top             =   60
               Width           =   1275
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   41
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   37
               Top             =   360
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpJtTxnLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepJtTxnLst
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
    
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  
  If adc("UoYN") = "Y" Then OptValWt(0).Value = True
  If adc("UoYN") = "N" Then OptValWt(1).Value = True
  'mw.131  assigning default value. showing always Show design=Yes
  adc("wShowDsg") = "Y"
  OptShowDsg(0).Value = True
  adc("UoValDivBy") = 1     'sv.30 setting default value
  
End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, ws_OrdBy As String
  Dim wTcCnd As String, wQtyStr As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  wTcCnd = IIF(adc("UoOmTc") <> "", " '" + adc("UoOmTc") + "' ", "")
  wTcCnd = wTcCnd + IIF(adc("UoOmTc1") <> "", IIF(wTcCnd <> "", ", ", "") + "'" + adc("UoOmTc1") + "' ", "")
  wTcCnd = wTcCnd + IIF(adc("UoOmTc2") <> "", IIF(wTcCnd <> "", ", ", "") + "'" + adc("UoOmTc2") + "' ", "")
  wTcCnd = IIF(wTcCnd <> "", " and OmTc In (" + wTcCnd + ") ", "")
    
  wCnd = wCnd + wTcCnd
  wCnd = IIF(wCnd = "", " where ", wCnd + " and ") + " (Select PValue From Param Where PTyp='TC' and PMCd=OmTc)='" + ctJtTcTyp + "' "

  If wTcCnd <> "" Then
    wQtyStr = IIF(adc("UoOmTc1") <> "", " (OmTc='" + adc("UoOmTc1") + "' and '" + adc("UoYn1") + "'='N') ", "")
    wQtyStr = wQtyStr + IIF(adc("UoOmTc2") <> "", IIF(wQtyStr <> "", " Or ", "") + " (OmTc='" + adc("UoOmTc2") + "' and '" + adc("UoYn2") + "'='N') ", "")

    wQtyStr = IIF(wQtyStr <> "", " (Case When " + wQtyStr + " Then -1 Else 1 End) ", "")
    wQtyStr = IIF(wQtyStr = "", "1", wQtyStr)
  Else
    wQtyStr = IIF(wQtyStr = "", "1", wQtyStr)
  End If
  
  Dim wValDivBy As Double
  'sv.30 when divied by amount is blank then default value is 1
  wValDivBy = IIF(adc("UoValDivBy") = 0, 1, adc("UoValDivBy"))

  '578.22 if currency code not given it takes base currency from head else takes from scope
  Dim ws_BaseCurCd As String, wb_BaseCur As Boolean
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_BaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_BaseCur = False
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If

  'sv.30 wValDivBy added to OdSalPrc, qMetVal, qDiaVal, qCsVal, qXVal
  '578.22 conversion factor added to OdSalPrc, qMetVal, qDiaVal, qCSVal, qXVal
  wSqlStrg = "Select OmCoCd,OmTc,OmYy,OmChr,OmNo,max(OmDt) as OmDt,max(OmFrDmLoc) as OmFrDmLoc,max(OmToDmLoc) as OmToDmLoc," + _
            "max(OmCmCd) as OmCmCd,max(OmPoNo) as OmPoNo,max(OmPoDt) as OmPoDt,OdSr,max(OdDmCd) as OdDmCd,max(OdSfx) as OdSfx," + _
            "max(OdDmSz) as OdDmSz,max(OdDmStkCoCd) as OdDmStkCoCd,max(OdDmStkYy) as OdDmStkYy, max(OdDmStkChr) as OdDmStkChr,max(OdDmStkNo) as OdDmStkNo," + _
            "max(OdDmCol) as OdDmCol,max(OdKt) as OdKt ,max(OdOrdQty)*" + wQtyStr + " as OdOrdQty, " + _
            "max(Convert(Decimal(14,3)," + IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + "/" + CStr(wValDivBy) + ")) as OdSalPrc, " + _
            "case when max(OdTc)='JRI' then max(OdRepGrWt * OdOrdQty) else sum((case when IsNull(OrRmCtg, '') in ('D','C') then IsNull(OrWt, 0)/5 else IsNull(OrWt, 0) end)* OdOrdQty) end *" + wQtyStr + "  as qGrsWt ," + _
            "sum((case when IsNull(OrRmCtg, '') in ('G','P','S','L','M')  then  IsNull(OrWt, 0) else 0 end)* OdOrdQty)*" + wQtyStr + "  as qMetWt ," + _
            "sum(((case when Isnull(OrRmCtg,'') in ('G', 'P', 'S', 'L', 'M') then RmPurityWt Else 0 end) /(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') " + _
            "then IsNull(qPurityWt, 1) Else 1 End))* (IsNull(OrWt, 0)*OdOrdQty))*" + wQtyStr + "  as qPureWt,sum((case when IsNull(OrRmCtg, '') ='D' then  IsNull(OrWt, 0) else 0 end) * OdOrdQty ) *" + wQtyStr + "  as qDiaWt," + _
            "sum((case when IsNull(OrRmCtg, '') ='D' then IsNull(OrQty,0) else 0 end)* OdOrdQty)*" + wQtyStr + "   as qDiaQty ," + _
            "sum((case when IsNull(OrRmCtg, '') ='C' then IsNull(OrWt, 0) else 0 end)* OdOrdQty)*" + wQtyStr + "   as qCsWt," + _
            "sum((case when IsNull(OrRmCtg, '') ='C' then IsNull(OrQty,0) else 0 end)* OdOrdQty)*" + wQtyStr + "   as qCsQty, " + _
            "sum((case when IsNull(OrRmCtg, '') ='X' then  IsNull(OrWt, 0) else 0 end)* OdOrdQty)*" + wQtyStr + "  as qXWt, " + _
            "sum((case when IsNull(OrRmCtg, '') in ('G','P','S','L','M')  then  IsNull(" + IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + "/" + CStr(wValDivBy) + ", 0) else 0 end)* OdOrdQty)*" + wQtyStr + "  as qMetVal ," + _
            "sum((case when IsNull(OrRmCtg, '') ='D' then IsNull(" + IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + "/" + CStr(wValDivBy) + ", 0) else 0 end) * OdOrdQty ) *" + wQtyStr + "  as qDiaVal," + _
            "sum((case when IsNull(OrRmCtg, '') ='C' then IsNull(" + IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + "/" + CStr(wValDivBy) + ", 0) else 0 end)* OdOrdQty)*" + wQtyStr + "   as qCsVal," + _
            "sum((case when IsNull(OrRmCtg, '') ='X' then IsNull(" + IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + "/" + CStr(wValDivBy) + ", 0) else 0 end)* OdOrdQty)*" + wQtyStr + "  as qXVal " + _
            "From OrdMst Join OrdDsg On OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + _
            "Left Outer Join OrdRm On OrCoCd=OdCoCd and OrTc=OdTc and OrYy=OdYy and OrChr=OdChr and OrNo=OdNo and OrSr=OdSr and OrOdIdNo=OdIdNo and OrPrtKey=OdPrtKey " + _
            "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= ''  and DmPrtKey=OdPrtKey " + _
            "Left Outer Join RmMst On RmCd=OrRmCd  and RmPrtKey=OrPrtKey " + _
            "Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on  RmPrtKey='C' and  RmCd= PValue " + _
            "where PTyp='RMCTG')PRm On qPMCd= OrRmCtg " + wCnd + " group by OmCoCd, OmTc, OmYy, OmChr, OmNo, OdSr " + _
            "Order by OmCoCd, OmTc, OmYy, OmChr, OmNo, OdSr "
            
            
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt
  
  
  Rep.TxtCoCdFr.SetText adc("UoCoCdFr"): Rep.TxtCoCdTo.SetText adc("UoCoCdTo"): Rep.TxtCoCdSel.SetText adc("UoCoCdSel")
    
  If wTcCnd = "" Then Rep.TxtTc.Suppress = True
  Rep.txtTcs.SetText IIF(adc("UoOmTc") <> "", "(" + adc("UoOmTc") + ")", "") + IIF(adc("UoOmTc1") <> "", IIF(adc("UoYn1") = "Y", " + ", " - ") + "(" + adc("UoOmTc1") + ")", "") _
                      + IIF(adc("UoOmTc2") <> "", IIF(adc("UoYn2") = "Y", " + ", " - ") + "(" + adc("UoOmTc2") + ")", "")
                      
  If adc("UoYN") = "Y" Then
    Rep.qGrsWt1.Suppress = True: Rep.g1GrsWt1.Suppress = True: Rep.gGrsWt1.Suppress = True
    Rep.qMetWt1.Suppress = True: Rep.g1MetWt1.Suppress = True: Rep.gMetWt1.Suppress = True
    Rep.qPureWt1.Suppress = True: Rep.g1PureWt1.Suppress = True: Rep.gPureWt1.Suppress = True
    Rep.qDiaWt1.Suppress = True: Rep.g1DWt1.Suppress = True: Rep.gDWt1.Suppress = True
    Rep.qCsWt1.Suppress = True: Rep.g1CWt1.Suppress = True: Rep.gCWt1.Suppress = True
    Rep.qXWt1.Suppress = True: Rep.g1xwt1.Suppress = True: Rep.gXWt1.Suppress = True
    Rep.TxtGrsWt.Suppress = True: Rep.TxtMetWt.Suppress = True: Rep.TxtPureWt.Suppress = True
    Rep.TxtDiaWt.Suppress = True: Rep.TxtCsWt.Suppress = True: Rep.TxtXWt.Suppress = True
    Rep.TxtDiaQty.Suppress = True:  Rep.TxtCsQty.Suppress = True
    Rep.qDiaQty1.Suppress = True: Rep.g1DQty1.Suppress = True:  Rep.gDQty1.Suppress = True
    Rep.qCsQty1.Suppress = True:  Rep.g1CQty1.Suppress = True:  Rep.gCQty1.Suppress = True
    'mw.131 if show value is true then suppress wt fields
    Rep.g1GrsWt2.Suppress = True: Rep.g1MetWt2.Suppress = True
    Rep.g1DWt2.Suppress = True:   Rep.g1CWt2.Suppress = True
    Rep.g1xwt2.Suppress = True:   Rep.g1PureWt2.Suppress = True
    Rep.g1DQty2.Suppress = True:  Rep.g1CQty2.Suppress = True
       
  ElseIf adc("UoYN") = "N" Then
     Rep.wSalVal1.Suppress = True: Rep.g1SalVal1.Suppress = True: Rep.gSalVal1.Suppress = True
     Rep.OdSalPrc1.Suppress = True: Rep.g1SalPrc1.Suppress = True: Rep.gSalPrc1.Suppress = True
     Rep.hSalPrc.Suppress = True: Rep.hSalVal.Suppress = True
     Rep.qMetVal1.Suppress = True:  Rep.qDiaVal1.Suppress = True
     Rep.qCsVal1.Suppress = True:   Rep.qXVal1.Suppress = True
     Rep.TxtMetVal.Suppress = True: Rep.TxtDiaVal.Suppress = True
     Rep.TxtCsVal.Suppress = True:  Rep.TxtXVal.Suppress = True
     Rep.g1MetVal1.Suppress = True:   Rep.g1DVal1.Suppress = True
     Rep.g1CVal1.Suppress = True:  Rep.g1XVal1.Suppress = True
     Rep.gMetVal1.Suppress = True:    Rep.gDVal1.Suppress = True
     Rep.gCVal1.Suppress = True:  Rep.gXVal1.Suppress = True
     'mw.131 if show wt is true then suppress value fields
     Rep.g1SalVal2.Suppress = True: Rep.g1SalPrc2.Suppress = True
     Rep.g1MetVal2.Suppress = True: Rep.g1DVal2.Suppress = True
     Rep.g1CVal2.Suppress = True:   Rep.g1XVal2.Suppress = True
     
  End If
  
  
  'mw.131 if show design=No then it suppress design fields else it show voucher details
  If adc("wShowDsg") = "N" Then
    Rep.Detsec.Suppress = True:     Rep.GrpHdSec.Suppress = True
    Rep.GrpFtSec1.Suppress = True:  Rep.LineRptFt.Suppress = False
    Rep.hSr.Suppress = True
    Rep.hDsgCd.Suppress = True:     Rep.hSfx.Suppress = True
    Rep.hSz.Suppress = True:        Rep.hStkNum.Suppress = True
    Rep.hColor.Suppress = True:     Rep.hKt.Suppress = True
    Rep.hSalPrc.Suppress = True
  Else
    If adc("UoYN") = "Y" Then
      Rep.hSalPrc.Suppress = False
    Else
      Rep.hSalPrc.Suppress = True
    End If
    Rep.GrpFtSec2.Suppress = True
    Rep.hVchrNo.Suppress = True:  Rep.hDt.Suppress = True
    Rep.hCmCd.Suppress = True:    Rep.hFrLoc.Suppress = True
    Rep.hToLoc.Suppress = True
  End If
  'sv.30 passing divide by value to dsr
  GRP_REP.SetFormula Rep, "wDivByVal", adc("UoValDivBy")
  
  'sv.30 when divied by amount>1 then follwing text will display in header
  If adc("UoValDivBy") > 1 Then Rep.TxtDivByVal.SetText "(Values in " + Format(adc("UoValDivBy"), "##,##,##0") + ")"
     
      
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc)
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin
End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
End Sub
Private Sub ADC_Load()
    Set moCn = adc.Connection
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
    Call InitProp(Me)
            
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  'sv.30 divided by amount cannot be 0
  Select Case UCase(IdName)
  Case Is = UCase("UoValDivBy")
    If pv_NewValue = 0 Then Cancel = True: ErrMsg = "Divided By Amount Cannot be Zero": Exit Sub
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_OmTcTyp = "JT"
  SetProp Me, IdName, When
End Sub
'mw.131 store ShowDesign Y/N value as per user selection
Private Sub OptShowDsg_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    adc("wShowDsg") = "Y"
  Case Is = 1
    adc("wShowDsg") = "N"
  End Select
End Sub
'mw.131 tooltip for option buttons
Private Sub OptShowDsg_GotFocus(Index As Integer)
  DispMsg FraOptShowDsg.ToolTipText, etInfo
End Sub

Private Sub OptValWt_GotFocus(Index As Integer)
  DispMsg FraShowValWt.ToolTipText, etInfo
   
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UoUsrCd") = adc.UsrCd
    adc("UoMnuCd") = adc.MenuCd
   
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm

  Rep.wCoCdLogo.SetText gs_CoNm
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If

  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

'  Rep.TxtCdCtg.SetText adc("UoCdCtg"):  Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
'  Rep.TxtFrCertDt.SetText adc("UoDtFr"):   Rep.TxtToCertDt.SetText adc("UoDtTo")
'  Rep.TxtFrCertNo.SetText adc("UoCdCertNoFr"):  Rep.TxtToCertNo.SetText adc("UoCdCertNoTo")
'  Rep.TxtFrRmLotNo.SetText adc("UoLotNoFr"):  Rep.TxtToRmLotNo.SetText adc("UoLotNoTo")
'  Rep.TxtFrBCoCd.SetText adc("UoCoCdFr"):  Rep.TxtFrBYy.SetText adc("UoBYyFr")
'  Rep.TxtFrBChr.SetText adc("UoBChrFr"):  Rep.TxtFrBNo.SetText adc("UoBNoFr")
'  Rep.TxtToBCoCd.SetText adc("UoCoCdTo"):  Rep.TxtToBYy.SetText adc("UoBYyTo")
'  Rep.TxtToBChr.SetText adc("UoBChrTo"):  Rep.TxtToBagNo.SetText adc("UoBNoTo")
'  Rep.TxtFrRmCode.SetText adc("UoRmCdFr"):  Rep.TxtToRmCode.SetText adc("UoRmCdTo")
'  Rep.TxtFrFMBarCd.SetText adc("UoCdFMBarCdFr"):  Rep.TxtToFMBarCd.SetText adc("UoCdFMBarCdTo")

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
    Set moCn = Nothing
    Set Rep = Nothing
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 1
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  End Select
End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

Private Sub OptValWt_Click(Index As Integer)
'sv.30 when Wt option is selected then divied by amount will be 1 and disabled otherwise enabled
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
  adc("UoValDivBy").Enabled = True
Case Is = 1
  adc("UoYN") = "N"
  adc("UoValDivBy") = 1
  adc("UoValDivBy").Enabled = False
End Select
End Sub


