VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPpcOrdPln 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Planning "
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   37
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10950
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4140
      TabIndex        =   45
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   46
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
         TabIndex        =   47
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
      Left            =   5340
      TabIndex        =   44
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   41
         Top             =   -165
         Visible         =   0   'False
         Width           =   495
         lastProp        =   600
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
      Height          =   9645
      Left            =   0
      TabIndex        =   40
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   39
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPpcOrdPln.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPpcOrdPln.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   43
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   38
               Top             =   360
               Width           =   9585
               _ExtentX        =   16907
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   42
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptPcsLoad 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   2040
               TabIndex        =   70
               ToolTipText     =   "Select If to Show Pieces or Load"
               Top             =   3930
               Width           =   1305
               Begin VB.OptionButton OptPcsLoad 
                  Caption         =   "Pieces"
                  BeginProperty Font 
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
                  TabIndex        =   33
                  ToolTipText     =   "Select If to Show Pieces or Load"
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton OptPcsLoad 
                  Caption         =   "Load"
                  BeginProperty Font 
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
                  Left            =   0
                  TabIndex        =   34
                  ToolTipText     =   "Select If to Show Pieces or Load"
                  Top             =   285
                  Width           =   1095
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2040
               TabIndex        =   1
               ToolTipText     =   "Enter Prcoess Stage"
               Top             =   765
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4260
               TabIndex        =   19
               ToolTipText     =   "Enter To Date"
               Top             =   2190
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               ReCalcOn        =   "UODTFR"
               IdName          =   "UODTTO"
               CmpStr          =   "OpdPlnDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2040
               TabIndex        =   18
               ToolTipText     =   "Enter From Date"
               Top             =   2190
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               ReCalcParent    =   "UODTTO"
               IdName          =   "UODTFR"
               CmpStr          =   "OpdPlnDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4260
               TabIndex        =   31
               ToolTipText     =   "Enter To Line"
               Top             =   3615
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOLINETO"
               IdName          =   "UOLINETO"
               CmpStr          =   "OpLine<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   6480
               TabIndex        =   29
               ToolTipText     =   "EnterDesign Code Selection"
               Top             =   3330
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2040
               TabIndex        =   30
               ToolTipText     =   "Enter From Line"
               Top             =   3615
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOLINEFR"
               IdName          =   "UOLINEFR"
               CmpStr          =   "OpLine>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5055
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Year"
               Top             =   1335
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2835
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Year"
               Top             =   1335
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   2040
               TabIndex        =   5
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1335
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   3255
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Character"
               Top             =   1335
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4260
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1335
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5475
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Character"
               Top             =   1335
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   2040
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Number"
               Top             =   1620
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4260
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Number"
               Top             =   1620
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5190
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1620
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   2970
               TabIndex        =   12
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1620
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   2040
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   480
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OpdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4260
               TabIndex        =   3
               ToolTipText     =   "Enter To Customer Code"
               Top             =   1050
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
               Index           =   8
               Left            =   2040
               TabIndex        =   2
               ToolTipText     =   "Enter From Customer Code"
               Top             =   1050
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4260
               TabIndex        =   16
               ToolTipText     =   "Enter To Priority Code"
               Top             =   1905
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
               CmpStr          =   "OdPrtCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4260
               TabIndex        =   21
               ToolTipText     =   "Enter To Delivery Date"
               Top             =   2475
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2040
               TabIndex        =   20
               ToolTipText     =   "Enter From Delivery Date"
               Top             =   2475
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4260
               TabIndex        =   25
               ToolTipText     =   "Enter To Design Category To"
               Top             =   3045
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   2040
               TabIndex        =   24
               ToolTipText     =   "Enter From Design Category"
               Top             =   3045
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   4260
               TabIndex        =   28
               ToolTipText     =   "Enter To Design Code"
               Top             =   3330
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2040
               TabIndex        =   27
               ToolTipText     =   "Enter From Design Code"
               Top             =   3330
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   503
               MaxLength       =   15
               DataType        =   4
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   3510
               TabIndex        =   32
               ToolTipText     =   "Show Pieces /Load ('Y' For Pieces/ 'N' For Load)"
               Top             =   3900
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   6450
               TabIndex        =   4
               ToolTipText     =   "Ente Customer Selection"
               Top             =   1050
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCD In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6480
               TabIndex        =   17
               ToolTipText     =   "Ente Priority Code Selection"
               Top             =   1905
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   2040
               TabIndex        =   15
               ToolTipText     =   "Enter From Priority Code"
               Top             =   1905
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
               CmpStr          =   "OdPrtCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   6480
               TabIndex        =   26
               ToolTipText     =   "Enter Line Selection"
               Top             =   3045
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   4260
               TabIndex        =   23
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   2760
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   2040
               TabIndex        =   22
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   2760
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdExpDelDt >="
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
               Index           =   20
               Left            =   5100
               TabIndex        =   69
               Top             =   1620
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
               Index           =   19
               Left            =   5400
               TabIndex        =   68
               Top             =   1335
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
               Index           =   18
               Left            =   4980
               TabIndex        =   67
               Top             =   1335
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
               Index           =   16
               Left            =   2880
               TabIndex        =   66
               Top             =   1620
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
               Index           =   14
               Left            =   3180
               TabIndex        =   65
               Top             =   1335
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
               Index           =   13
               Left            =   2760
               TabIndex        =   64
               Top             =   1335
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Date"
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
               TabIndex        =   63
               Top             =   2760
               Width           =   2025
            End
            Begin VB.Label ALBL 
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
               Index           =   11
               Left            =   90
               TabIndex        =   62
               Top             =   3900
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Cd"
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
               TabIndex        =   61
               Top             =   3330
               Width           =   2025
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   8
               Left            =   90
               TabIndex        =   60
               ToolTipText     =   " --------------------------------------------------------"
               Top             =   3045
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Del Date"
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
               TabIndex        =   59
               Top             =   2475
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Priority Code"
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
               TabIndex        =   58
               Top             =   1905
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Cd"
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
               TabIndex        =   57
               Top             =   1050
               Width           =   2025
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
               Left            =   90
               TabIndex        =   56
               Top             =   480
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Tc/Yy/Chr"
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
               TabIndex        =   55
               Top             =   1335
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order No/Sr"
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
               TabIndex        =   54
               Top             =   1620
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Production Stage"
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
               Top             =   765
               Width           =   2025
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Plan Date"
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
               Left            =   90
               TabIndex        =   52
               Top             =   2190
               Width           =   2025
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   4
               Left            =   90
               TabIndex        =   51
               Top             =   3615
               Width           =   2025
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
               Left            =   6450
               TabIndex        =   50
               Top             =   30
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
               Left            =   4200
               TabIndex        =   49
               Top             =   30
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
               Left            =   2040
               TabIndex        =   48
               Top             =   30
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPpcOrdPln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepPpcOrdPln
Dim moCn As MwfLib.MDOConnection     '*** (Bef speed) bef 24/11/06
Dim wPpcPrc As String

Private Sub SetGroupSort()
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "max(OmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end ) From CustMst Where CmCtg= 'C' and CmCd= max(OmCmCd)) "
  GRP_REP.Add "OrderNo", "max(OdTc+OdYy+OdChr+str(OdNo))", "max(OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo)))", "", "", "Customer Code,Customer PO", "", "", "'('+max(OmCmCd)+ '/'+ max(OmPONo)+ ')' "
  GRP_REP.Add "Priority", "max(OdPrtCd)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= max(OdPrtCd)) "
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg))"
  GRP_REP.Add "Design Code", "max(OdDmCd)", "", "", "", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where DmTcTyp= 'DM' and DmCd= max(OdDmCd) and DmSz= '') "
  '6.1 Ord plan date, Exp Del date
  GRP_REP.Add "Ord Plan Date", "convert(varchar(8), max(OpdPlnDt), 112)", "convert(varchar, max(OpdPlnDt), " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Exp Del Date", "convert(varchar(8), max(OdExpDelDt), 112)", "convert(varchar, max(OdExpDelDt), " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  GRP_REP.Add "Line", "max(OpLine)", "", "", "", "", "", "", "(Select vPDesc from vParam where vPCoCd= max(OpCoCd) and vPTyp= 'LINE' and vPmCd = max(OpLine))"
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'**** Before 090205 ****'
  ' Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  ' space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  ' space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  ' OdTc , OdYy, OdChr, OdNo, OdSr, OdDmCd, OmCmCd,
  ' 0.0 as qQty1, 0.0 as qQty2, 0.0 as qQty3, 0.0 as qQty4,0.0 as qQty5,
  ' 0.0 as qQty6,   0.0 as qQty7, 0.0 as qQty8, 0.0 as qQty9, 0.0 as qQty10,
  ' 0.0 as qQty11, 0.0 as qQty12, 0.0 as qQty13, 0.0 as qQty14, 0.0 as qQty15,
  ' 0.0 as qLd1, 0.0 as qLd2, 0.0 as qLd3, 0.0 as qLd4,0.0 as qLd5,
  ' 0.0 as qLd6,   0.0 as qLd7, 0.0 as qLd8, 0.0 as qLd9, 0.0 as qLd10,
  ' 0.0 as qLd11, 0.0 as qLd12, 0.0 as qLd13, 0.0 as qLd14, 0.0 as qLd15,
  ' 0.0 as qCapacity
  ' from OrdDsg, OrdMst, OrdPln, OrdPlnDt where 1=2
'**** Before 090205 ****'

  '********* Zubin(090205) *********'
  ' Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  ' space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  ' space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  ' OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, OdDmCd as DmCd, OmCmCd as CmCd,
  ' 0.0 as qQty1, 0.0 as qQty2, 0.0 as qQty3, 0.0 as qQty4,0.0 as qQty5,
  ' 0.0 as qQty6, 0.0 as qQty7, 0.0 as qQty8, 0.0 as qQty9, 0.0 as qQty10,
  ' 0.0 as qQty11, 0.0 as qQty12, 0.0 as qQty13, 0.0 as qQty14, 0.0 as qQty15,
  ' 0.0 as qLd1, 0.0 as qLd2, 0.0 as qLd3, 0.0 as qLd4,0.0 as qLd5,
  ' 0.0 as qLd6, 0.0 as qLd7, 0.0 as qLd8, 0.0 as qLd9, 0.0 as qLd10,
  ' 0.0 as qLd11, 0.0 as qLd12, 0.0 as qLd13, 0.0 as qLd14, 0.0 as qLd15,
  ' 0.0 as qCapacity
  ' from OrdPlnDt, OrdPln, OrdDsg, OrdMst, DsgMst where 1=2
  '********* Zubin(090205) *********'
  '*** Report Sql ***
  
  Dim grpflds As String, wCnd As String, wSqlStrg As String   '*** Variables Displayed for the  Group Fields and Conditions and Sql String
  Dim CurrDate As Date
  Dim wStr As String
  Dim wStrPrc As String
  Dim ws_TotCapCnd As String
  
  Dim ws_PrdStg As String   '*** (Jen 2.14 Next)
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst

  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  
  '*** (Jen 2.14 Next) (changed the fields WSet, Cas, Fin, HSet, Ei to P1, P2, P3, P4, P5) ***
  ws_PrdStg = Trim(adc("UoPmCdFr"))
  wStr = "Opd" + ws_PrdStg: wStrPrc = "Dld" + ws_PrdStg: wPpcPrc = "Pc" + ws_PrdStg
  '*** (Jen 2.14 Next) (changed the fields WSet, Cas, Fin, HSet, Ei to P1, P2, P3, P4, P5) ***
  
  '*** (Bef 2.14 Next)
  'Select Case UCase(ADC("UoPmCdFr"))
  'Case Is = "WSET"
  '  wStr = "OpdwSet": wStrPrc = "DldwSet": wPpcPrc = "PcwSet"
  'Case Is = "CAS"
  '  wStr = "OpdCas": wStrPrc = "DldCas": wPpcPrc = "PcCas"
  'Case Is = "FIN"
  '  wStr = "OpdFin": wStrPrc = "DldFin": wPpcPrc = "PcFin"
  'Case Is = "HSET"
  '  wStr = "OpdhSet": wStrPrc = "DldhSet": wPpcPrc = "PchSet"
  'Case Is = "EI"
  '  wStr = "OpdEi": wStrPrc = "DldEi": wPpcPrc = "PcEi"
  'End Select
  '*** (Bef 2.14 Next)
   
  If adc("UoYN") = "Y" Then
    Rep.TxtHead.SetText ("Department Calendar For Pieces")
  Else
    Rep.TxtHead.SetText ("Department Calendar For Load")
  End If
  wCnd = adc.RepCond
  wCnd = IIF(wCnd <> "", " and " + wCnd, "")
  ws_TotCapCnd = IIF(adc("UoLineFr") = "", " ", " and PcLine>= '" + adc("UoLineFr") + "' ")
  ws_TotCapCnd = ws_TotCapCnd + IIF(adc("UoLineTo") = "", " ", " and PcLine<= '" + adc("UoLineTo") + "' ")

  '****** Sachin 3.02 - If fields in Join
  '6.1
  wSqlStrg = "Set dateformat " + DtFmtStr() + " " + _
             "Select " + grpflds + ", OpdTc, OpdYy, OpdChr, OpdNo, OpdSr, max(OdDmCd) as DmCd, " + _
             "max(OmCmCd) as CmCd, " + wFldLst(wStr, wStrPrc) + ", " + _
             CStr(moCn.GetFldVal("(select sum(" + wPpcPrc + ") from PpcCap where PcCoCd= '" + gs_CoCd + "' " + _
             "and PcLine<> '" + ctAllLineCd + "' " + ws_TotCapCnd + ") ")) + " as qCapacity " + _
             "from OrdPlnDt join OrdPln on OpCoCd= OpdCoCd and OpTc= OpdTc and OpYy= OpdYy " + _
             "    and OpChr= OpdChr and OpNo= OpdNo and OpSr= OpdSr " + _
             "join OrdDsg on OdCoCd= OpCoCd and OdTc= OpTc and OdYy= OpYy and OdChr= OpChr " + _
             "    and OdNo= OpNo and OdSr= OpSr " + _
             "join OrdMst on /* OmIdNo=OdOmIdNo and */ OmCoCd= OdCoCd and OmTc= OpTc and OmYy= OpYy and OmChr= OpChr " + _
             "    and OmNo= OpNo " + _
             "join DsgMst on /* DmIdNo=OdDmIdNo and */ DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
             "where OpdCoCd= '" + gs_CoCd + "' " + wCnd + " and " + wStr + "<> 0 " + _
             "group by OpdTc, OpdYy, OpdChr, OpdNo, OpdSr "
      
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub
Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoCmCdFr")
  Call InitProp(Me)
  Call SetGroupSort
End Sub
Private Sub ADC_KeyWhen()
  adc("UoUsrCd") = adc.UsrCd
  adc("UoMnuCd") = adc.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  If adc("UoDtFr") = "" Then adc("UoDtFr") = moCn.SrvrDate
  If adc("UoDelDtFr") = "" Then adc("UoDelDtFr") = "01/01/80"
  If adc("UoDelDtTo") = "" Then adc("UoDelDtTo") = "01/01/80"
  If adc("UoDelDt1Fr") = "" Then adc("UoDelDt1Fr") = "01/01/80"
  If adc("UoDelDt1To") = "" Then adc("UoDelDt1To") = "01/01/80"
  
  ' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added
  If adc("UoYN") = "Y" Then OptPcsLoad(0).Value = True
  If adc("UoYN") = "N" Then OptPcsLoad(1).Value = True
  ' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added

End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case IdName
  Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
    gs_CmCtg = "C"
  Case Is = UCase("UoPmCdFr")
    gs_PTyp = "PRDSTG"
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case IdName
  Case Is = UCase("UoCmCdFr"), UCase("UoCmCdTo"), UCase("UoCmCdSel")
    gs_CmCtg = "C"
  Case Is = UCase("UoPmCdFr")
    gs_PTyp = "PRDSTG"
  Case Is = UCase("UoDtTo")
    Cancel = True
  End Select
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case IdName
  Case Is = UCase("UoDtTo")
    If adc.Mode = xNorm Then
      'If ADC("UoDtFr") <> "" Then ADC("UoDtTo") = DateAdd("D", 14, CDate(ADC("UoDtFr")))
      If adc("UoDtFr") <> "" Then adc("UoDtTo") = DateAdd("D", 14, adc("UoDtFr"))
    End If
  End Select
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

' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added
Private Sub OptPcsLoad_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added
Private Sub OptPcsLoad_GotFocus(Index As Integer)
  DispMsg FraOptPcsLoad.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 18/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText gs_CoNm
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
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
Private Function wFldLst(ms_FldName As String, ms_FldNmLoad As String) As String
  Dim ws_TempStrQty, ws_TempStrVal, wSqlStrg As String
  Dim ws_CurDate As String, wi_Ctr As Integer
  
  If adc("UoDtFr") <> "" Then
    ws_CurDate = adc("UoDtFr")
    wi_Ctr = 0
    Do While CDate(ws_CurDate) <= CDate(adc("UoDtTo"))
      wi_Ctr = wi_Ctr + 1
      ws_TempStrQty = ws_TempStrQty + IIF(ws_TempStrQty = "", "", ", ") + " Sum(case when OpdPlnDt= '" + ws_CurDate + "' then " + ms_FldName + " else 0 end) as qQty" + CStr(wi_Ctr)
      ws_TempStrVal = ws_TempStrVal + IIF(ws_TempStrVal = "", "", ", ") + " Sum(case when OpdPlnDt= '" + ws_CurDate + "' then " + ms_FldName + " else 0 end) * " + _
                      "IsNull((select " + ms_FldNmLoad + " from DsgLoad where DldDmCd= max(OdDmCd) and DldWH= max(OdWH)), " + _
                      " IsNull((select " + ms_FldNmLoad + " from DsgLoad " + _
                      " where DldPrdCtg= (Select DmPrdCtg from DsgMst where DmTcTyp= 'DM' and " + _
                      " DmCd= max(OdDmCd) and DmSz= '') " + _
                      " and DldDmCd= '' and DldWH= max(OdWH)), 0)) as qLd" + CStr(wi_Ctr)
      
      GRP_REP.SetFormula Rep, "hDt" + CStr(wi_Ctr), "'" + ws_CurDate + "'"
      
      If adc("UoYN") = "N" Then
        GRP_REP.SetFormula Rep, "wVal" + CStr(wi_Ctr), "({rdo." + "qLd" + CStr(wi_Ctr) + "})"
      Else
        GRP_REP.SetFormula Rep, "wVal" + CStr(wi_Ctr), "({rdo." + "qQty" + CStr(wi_Ctr) + "})"
      End If
      ws_CurDate = DateAdd("D", 1, ws_CurDate)
    Loop
    If ws_TempStrVal <> "" Then
      wFldLst = ws_TempStrQty + ", " + ws_TempStrVal
    Else
      wFldLst = ws_TempStrQty
    End If
    GRP_REP.SetFormula Rep, "wFrLine", "'" + adc("UoLineFr") + "'"
    GRP_REP.SetFormula Rep, "wToLine", "'" + adc("UoLineTo") + "'"
    'GRP_REP.SetFormula Rep, "wPrdStg", "'" + adc("UoPmCdFr") + "'"     '*** (Bef 2.14 Next)
    '*** (Bef 2.14 Next)
    GRP_REP.SetFormula Rep, "wPrdStg", "'" + moCn.GetFldVal("Select PValue From Param where PTyp= 'PRDSTG' and PMCd= '" + adc("UoPmCdFr") + "' ") + "'"
    '*** (Bef 2.14 Next)
    
    GRP_REP.SetFormula Rep, "wPieceQuantity", "'" + adc("UoYN") + "'"
  End If
End Function
