VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpCustmPurLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Purchase Listing"
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
      TabIndex        =   33
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
      TabIndex        =   34
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
         TabIndex        =   39
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
      TabIndex        =   35
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
      TabIndex        =   38
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   37
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpCustmPurLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpCustmPurLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
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
               TabIndex        =   36
               Top             =   360
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   40
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptVch 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1515
               TabIndex        =   67
               ToolTipText     =   "Show Open Or Closed Or Both Vouchers"
               Top             =   3375
               Width           =   4395
               Begin VB.OptionButton OptVch 
                  Caption         =   "Closed"
                  BeginProperty Font 
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
                  Left            =   0
                  TabIndex        =   70
                  ToolTipText     =   "Show Open Or Closed Or Both Vouchers"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptVch 
                  Caption         =   "Open"
                  BeginProperty Font 
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
                  Left            =   1260
                  TabIndex        =   69
                  ToolTipText     =   "Show Open Or Closed Or Both Vouchers"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptVch 
                  Caption         =   "Both"
                  BeginProperty Font 
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
                  Left            =   2550
                  TabIndex        =   68
                  ToolTipText     =   "Show Open Or Closed Or Both Vouchers"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4005
               TabIndex        =   12
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   1095
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "TDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   1095
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "TDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4785
               TabIndex        =   5
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "TzYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2295
               TabIndex        =   2
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   525
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "TzYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   1
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "TzTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2715
               TabIndex        =   3
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "TzChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4005
               TabIndex        =   4
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "TzTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5205
               TabIndex        =   6
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   525
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "TzChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   4905
               TabIndex        =   10
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
               CmpStr          =   "TzSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2415
               TabIndex        =   8
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   810
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
               CmpStr          =   "TzSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
               CmpStr          =   "TzNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4005
               TabIndex        =   9
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   810
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
               CmpStr          =   "TzNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4005
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1950
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "TzRmZCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6510
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1950
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "TzRmZCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4005
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "RmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   6510
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1665
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1515
               TabIndex        =   13
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1380
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1515
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1515
               TabIndex        =   19
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1950
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "TzRmZCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4005
               TabIndex        =   14
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   1380
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "RmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   6510
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   1380
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "RmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4005
               TabIndex        =   23
               ToolTipText     =   "Enter To Supplier"
               Top             =   2235
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "TSuppCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   6510
               TabIndex        =   24
               ToolTipText     =   "Enter Supplier Selection"
               Top             =   2235
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "TSuppCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   1515
               TabIndex        =   22
               ToolTipText     =   "Enter From Supplier"
               Top             =   2235
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "TSuppCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4005
               TabIndex        =   26
               ToolTipText     =   "Enter To Bill Of Entry Number "
               Top             =   2520
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOETO"
               IdName          =   "UOBOETO"
               CmpStr          =   "TBOENo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6510
               TabIndex        =   27
               ToolTipText     =   "Enter Bill Of Entry Number Selection"
               Top             =   2520
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBOESEL"
               IdName          =   "UOBOESEL"
               CmpStr          =   "TBOENo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1515
               TabIndex        =   25
               ToolTipText     =   "Enter From Bill Of Entry Number "
               Top             =   2520
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOBOEFR"
               IdName          =   "UOBOEFR"
               CmpStr          =   "TBOENo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   240
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TzCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   1515
               TabIndex        =   31
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   3090
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
               CmpStr          =   "TCurrCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   1515
               TabIndex        =   32
               ToolTipText     =   "Show Closed (Yes / No) or ' ' for Both"
               Top             =   3375
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
               CmpStr          =   "TzCls = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1515
               TabIndex        =   28
               ToolTipText     =   "Enter Improt Type From "
               Top             =   2805
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "TImpTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4005
               TabIndex        =   29
               ToolTipText     =   "Enter Improt Type To"
               Top             =   2805
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "TImpTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   6510
               TabIndex        =   30
               ToolTipText     =   "Enter Improt Type Selection"
               Top             =   2805
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "TImpTyp in"
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Import Typ"
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
               Left            =   60
               TabIndex        =   66
               ToolTipText     =   "Location"
               Top             =   2790
               Width           =   1485
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Closed"
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
               Left            =   60
               TabIndex        =   65
               ToolTipText     =   "Location"
               Top             =   3360
               Width           =   1485
            End
            Begin VB.Label LblCurrCd 
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   60
               TabIndex        =   64
               ToolTipText     =   "Location"
               Top             =   3075
               Width           =   1485
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
               Index           =   0
               Left            =   60
               TabIndex        =   63
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label LblBOENo 
               BackStyle       =   0  'Transparent
               Caption         =   "BOE No"
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
               Left            =   60
               TabIndex        =   62
               Top             =   2520
               Width           =   1425
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
               Left            =   60
               TabIndex        =   61
               Top             =   2235
               Width           =   1425
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
               Left            =   60
               TabIndex        =   60
               ToolTipText     =   "Location"
               Top             =   1380
               Width           =   1425
            End
            Begin VB.Label LblRmSCtg 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm SubCtg"
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
               Left            =   60
               TabIndex        =   59
               Top             =   1665
               Width           =   1425
            End
            Begin VB.Label LblRmCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Code"
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
               Left            =   60
               TabIndex        =   58
               Top             =   1950
               Width           =   1425
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
               Index           =   11
               Left            =   60
               TabIndex        =   57
               Top             =   525
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch No/Sr"
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
               Left            =   60
               TabIndex        =   56
               Top             =   810
               Width           =   1605
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   28
               Left            =   5145
               TabIndex        =   55
               Top             =   525
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   29
               Left            =   2655
               TabIndex        =   54
               Top             =   525
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   32
               Left            =   4725
               TabIndex        =   53
               Top             =   525
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   33
               Left            =   2235
               TabIndex        =   52
               Top             =   525
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   34
               Left            =   4845
               TabIndex        =   51
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   35
               Left            =   2355
               TabIndex        =   50
               Top             =   810
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Date"
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
               Left            =   60
               TabIndex        =   49
               Top             =   1095
               Width           =   1425
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
               Left            =   6510
               TabIndex        =   48
               Top             =   0
               Width           =   1575
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
               Left            =   4005
               TabIndex        =   47
               Top             =   0
               Width           =   1275
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
               Left            =   1515
               TabIndex        =   46
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpCustmPurLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim Rep As New EmrRepCustmPurLst
Dim ws_MultiCoMod As Boolean
Private Sub SetGroupSort()
  GRP_REP.Add "Rm Ctg", "Rtrim(RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmCtg)+Rtrim(RmSCtg)", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
  GRP_REP.Add "Custm RmCd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "RmCtg,RmSCtg,RmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
  GRP_REP.Add "VoucherNo", "TTc+TYy+TChr+str(TNo)", "TTc+'/'+TYy+'/'+TChr+'/'+LTrim(Str(TNo))", "", "", "Vch date", "", "", ""
  '6.1
  GRP_REP.Add "Vch Date", "convert(char(8),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Supplier", "TSuppCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'S' and CmCd= TSuppCd) "
  GRP_REP.Add "BOE No", "TBOENo", "", "", "", "", "", "", "TBOEDt"
  GRP_REP.Add "Import Typ", "TImpTyp", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp= 'IMPTYP' and PMCd= TImpTyp) "
  GRP_REP.Add "Closed", "TzCls", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp= 'YN' and PMCd= TzCls) "
  '6.1
  GRP_REP.Add "Use Bef Dt", "convert(char(8),TDt,112)", "convert(varchar, TDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "", "", "", "", "", ""
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If adc("UoYN") = "Y" Then
    OptVch(0).Value = True
  ElseIf adc("UoYN") = "N" Then
    OptVch(1).Value = True
  ElseIf adc("UoYN") = "" Then
    OptVch(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'TCoCd, TTc, TYy, TChr, TNo, TDt, TSuppCd, TBillNo, TBillDt, TCnvRt,
  'TCurrCd, TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TTotAmt as qTTotAmt,
  'TzSr, TzRmZCd, TzRmCdDesc, TzPcPerCt, TzRmQty, TzRmWt, TzRmDespWt, TzCls,
  'TzPurRt As qTzPurRt, TzPurAmt As qTzPurAmt
  'From Txn, TxndZ, RmMst
  'Where 1 = 2
  '*** Report Sql ***
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wPurStr As String, wTTotAmtStr As String
  Dim ws_BaseCurCd As String
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New Fields In Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  '*** This Report should always have the detail section ***
  '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
    If wGrp(i) = "(None)" Then
      wNonePos = i: Exit For
    End If
  Next i
  
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TzRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGTO")) <> "", " TzRmCtg <= '" + Trim(adc("UORMCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGSEL")) <> "", " TzRmCtg In (" + Trim(adc("UORMCTGSEL")) + ")", "")

    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TzRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TzRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TzRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
  
  '*** If Detail option is not selected then report is not shown ***
  If wDetPos = 0 Then DispMsg "Detail Option Has To Be Selected For This Report", etError
  
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  If adc("UoCurCdFr") <> "" Then
    wPurStr = "TzPurRt As qTzPurRt, TzPurAmt As qTzPurAmt "
    wTTotAmtStr = "TNetAmt as qTNetAmt, TAddAmt1 as qTAddAmt1, TAddAmt2 as qTAddAmt2, TTotAmt as qTTotAmt, "
  Else
    wPurStr = " (case when TMulDiv= 'M' then TzPurRt/TCnvRt else TzPurRt*TCnvRt end) as qTzPurRt, (case when TMulDiv= 'M' then TzPurAmt/TCnvRt else TzPurAmt*TCnvRt end) qTzPurAmt "
    wTTotAmtStr = "(case when TMulDiv= 'M' then TNetAmt/TCnvRt Else TNetAmt/TCnvRt End) as qTNetAmt, " + _
                  " (case when TMulDiv= 'M' then TAddAmt1/TCnvRt Else TAddAmt1*TCnvRt End) as qTAddAmt1, " + _
                  " (case when TMulDiv= 'M' then TAddAmt2/TCnvRt Else TAddAmt2*TCnvRt End) as qTAddAmt2, " + _
                  " (case when TMulDiv= 'M' then TTotAmt/TCnvRt Else TTotAmt*TCnvRt End) as qTTotAmt, "
  End If
  '****** Sachin 3.02 - Id fields in Joins
  wSqlStrg = "Select " + grpflds + ", " + _
              "TCoCd, TTc, TYy, TChr, TNo, TDt, TSuppCd, TBillNo, TBillDt, TCnvRt, " + _
              "TCurrCd, " + wTTotAmtStr + _
              "TzSr, TzRmZCd, TzRmCdDesc, TzPcPerCt, TzRmQty, TzRmWt, TzRmDespWt, " + wPurStr + _
              "From Txn Join TxndZ on TzTIdNo=TIdNo and TzCoCd= TCoCd and TzTc= TTc and TzYy= TYy and TzChr= TChr and TzNo= TNo and TzSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " And TzPrtKey=TPrtKey ", "") + _
              "Join RmMst on RmCd= TzRmZCd and RmZ= 'Y' " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TzPrtKey ", "") + _
              "where " + IIF(gs_Partition = ctCurrPrtn, " TPrtKey='" + ctCurrPrtn + "' And ", "") + " (Select PValue from Param where PTyp= 'TC' and PMCd= TTc)= 'TZ' and TzTc >='TZ'" + wCnd + " " + _
              IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", "") + wAddlCnd, "") + _
              "Order By TTc, TYy, TChr, TNo, TzSr"
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
  If wDetPos = 1 Or wNonePos = 2 Then Rep.G1F.Suppress = True: Rep.G2F.Suppress = True
  If wDetPos = 2 Or wNonePos = 3 Then Rep.G1F.Suppress = True
  'If wDetPos = 3 Or wNonePos = 0 Then Rep.G1F.Suppress = True: Rep.G2F.Suppress = True
  
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TzRmZCd}", 7)
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.TzRmCdDesc}", 7)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub
'''' std code not to be changed
Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
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
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  Set adc.FirNKeyCtl = adc("UoTTcFr")
  Call InitProp(Me)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = True
  gb_RmZ = True
  gs_CmCtg = "S"
  gs_Tbl = "Param"
  gs_PTyp = "WIPVAL"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = False '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  gs_TxnTcTyp = "TZ"
  gs_PTyp = "IMPTYP"
  SetProp Me, IdName, When
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptVch_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
Case Is = 2
  adc("UoYN") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptVch_GotFocus(Index As Integer)
  DispMsg FraOptVch.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
  Call EnaDisaCmds(False)
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  Rep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  Rep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
  
  Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
  Set Rep = Nothing
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton)
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.FontBold = True
      pv_CmdBut.CausesValidation = True
    Else
      pv_CmdBut.BackColor = &H8000000F
      pv_CmdBut.FontBold = False
    End If
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
