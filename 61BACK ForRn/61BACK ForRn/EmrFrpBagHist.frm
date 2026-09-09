VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBagHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag History"
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
      Left            =   12270
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4965
      TabIndex        =   49
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   50
         Top             =   75
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
         TabIndex        =   51
         Top             =   90
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
      Left            =   6240
      TabIndex        =   48
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   45
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
      TabIndex        =   41
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
      TabIndex        =   44
      Top             =   60
      Width           =   15225
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   43
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBagHist.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBagHist.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   47
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   42
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   14865
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Complete History"
               BeginProperty Font 
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
               Left            =   30
               TabIndex        =   80
               Tag             =   "ShowFgMvm"
               ToolTipText     =   "Click To Show Complete History "
               Top             =   6600
               Visible         =   0   'False
               Width           =   2015
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Addl Dets"
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
               Index           =   1
               Left            =   30
               TabIndex        =   38
               Tag             =   "ShowFgMvm"
               ToolTipText     =   "Click To Show Stock Rate Field"
               Top             =   5640
               Width           =   2015
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Stk Rate"
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
               Index           =   0
               Left            =   30
               TabIndex        =   32
               Tag             =   "ShowFgMvm"
               ToolTipText     =   "Click To Show Stock Rate Field"
               Top             =   3720
               Width           =   1995
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1845
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   31
               Top             =   3450
               Width           =   2340
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show FG Move- ments Trans."
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
               Index           =   3
               Left            =   30
               TabIndex        =   37
               Tag             =   "ShowFgMvm"
               ToolTipText     =   "Click To Show FG Movement Transactions"
               Top             =   5160
               Width           =   2015
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4395
               TabIndex        =   26
               ToolTipText     =   "Enter To Current Bag Location"
               Top             =   2880
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   " "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6840
               TabIndex        =   27
               ToolTipText     =   "Enter Current Bag Location Selection"
               Top             =   2880
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1845
               TabIndex        =   25
               ToolTipText     =   "Enter From Current Bag Location"
               Top             =   2880
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5175
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Year"
               Top             =   1170
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2625
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Year"
               Top             =   1170
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1845
               TabIndex        =   1
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1170
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3045
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Character"
               Top             =   1170
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4395
               TabIndex        =   4
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1170
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5595
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Character"
               Top             =   1170
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5295
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1455
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2745
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1455
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1845
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Number"
               Top             =   1455
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4395
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Number"
               Top             =   1455
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4395
               TabIndex        =   18
               ToolTipText     =   "Enter To Bag Opening Date"
               Top             =   2025
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "BOpnDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   1845
               TabIndex        =   17
               ToolTipText     =   "Enter From Bag Opening Date"
               Top             =   2025
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "BOpnDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4395
               TabIndex        =   14
               ToolTipText     =   "Enter To Bag Year"
               Top             =   1740
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "BYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1845
               TabIndex        =   11
               ToolTipText     =   "Enter From Bag Year"
               Top             =   1740
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "BYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2265
               TabIndex        =   12
               ToolTipText     =   "Enter From Bag Character"
               Top             =   1740
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3045
               TabIndex        =   13
               ToolTipText     =   "Enter From Bag Number"
               Top             =   1740
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOFR"
               IdName          =   "UOBNOFR"
               CmpStr          =   "BNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   4815
               TabIndex        =   15
               ToolTipText     =   "Enter To Bag Character"
               Top             =   1740
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "BChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5595
               TabIndex        =   16
               ToolTipText     =   "Enter To Bag Number"
               Top             =   1740
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBNOTO"
               IdName          =   "UOBNOTO"
               CmpStr          =   "BNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4395
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Code"
               Top             =   2310
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "BOdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6840
               TabIndex        =   21
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2310
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "BOdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1845
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Code"
               Top             =   2310
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "BOdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4395
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer"
               Top             =   2595
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
               Left            =   6840
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2595
               Width           =   7950
               _ExtentX        =   14023
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
               Left            =   1845
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer"
               Top             =   2595
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
               Index           =   8
               Left            =   4395
               TabIndex        =   29
               ToolTipText     =   "Enter To Worker"
               Top             =   3165
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKTO"
               IdName          =   "UOWRKTO"
               CmpStr          =   "BWrk <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   6840
               TabIndex        =   30
               ToolTipText     =   "Enter Worker Selection"
               Top             =   3165
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOWRKSEL"
               IdName          =   "UOWRKSEL"
               CmpStr          =   "BWrk In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1845
               TabIndex        =   28
               ToolTipText     =   "Enter From Worker"
               Top             =   3165
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOWRKFR"
               IdName          =   "UOWRKFR"
               CmpStr          =   "BWrk >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1845
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   885
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   2415
               TabIndex        =   36
               ToolTipText     =   "Enter 'Y' To Show Fg Movement Transactions and 'N' Not To Show"
               Top             =   5295
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1845
               TabIndex        =   34
               ToolTipText     =   "Specify The field On Which The Rm Summary Subreport Has To Be Grouped"
               Top             =   4290
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   1845
               TabIndex        =   35
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   4830
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2415
               TabIndex        =   33
               ToolTipText     =   "Enter 'Y' To Show Stock rate"
               Top             =   3855
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   2415
               TabIndex        =   77
               ToolTipText     =   "Show Open Bags (Yes / No / All)"
               Top             =   5760
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
               Index           =   16
               Left            =   1845
               TabIndex        =   78
               ToolTipText     =   "Show 3D Picture Or Hand Sketch ? (Blank For No Picture)"
               Top             =   6240
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2400
               TabIndex        =   81
               ToolTipText     =   "Click To Show Complete History "
               Top             =   6600
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Pic (3D or HSk)"
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
               TabIndex        =   79
               ToolTipText     =   "Location"
               Top             =   6240
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter SEO Pwd"
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
               Left            =   60
               TabIndex        =   76
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Bag Typ"
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
               TabIndex        =   75
               ToolTipText     =   "Location"
               Top             =   4830
               Width           =   1395
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
               Index           =   11
               Left            =   60
               TabIndex        =   74
               Top             =   885
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Report Grouping"
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
               Height          =   555
               Index           =   8
               Left            =   60
               TabIndex        =   73
               Top             =   4290
               Width           =   1665
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Worker"
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
               Left            =   60
               TabIndex        =   72
               Top             =   3165
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Opn Date"
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
               Left            =   60
               TabIndex        =   71
               Top             =   2025
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Yy/Chr/No"
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
               Left            =   60
               TabIndex        =   70
               Top             =   1740
               Width           =   1605
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
               Left            =   2205
               TabIndex        =   69
               Top             =   1740
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
               Left            =   4755
               TabIndex        =   68
               Top             =   1740
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
               Index           =   20
               Left            =   2985
               TabIndex        =   67
               Top             =   1740
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
               Index           =   21
               Left            =   5535
               TabIndex        =   66
               Top             =   1740
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Cd"
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
               Left            =   60
               TabIndex        =   65
               Top             =   2310
               Width           =   1425
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
               Left            =   60
               TabIndex        =   64
               Top             =   2595
               Width           =   1425
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
               Index           =   31
               Left            =   2685
               TabIndex        =   63
               Top             =   1455
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
               Index           =   30
               Left            =   5235
               TabIndex        =   62
               Top             =   1455
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
               Index           =   27
               Left            =   2565
               TabIndex        =   61
               Top             =   1170
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
               Index           =   26
               Left            =   5115
               TabIndex        =   60
               Top             =   1170
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
               Index           =   25
               Left            =   2985
               TabIndex        =   59
               Top             =   1170
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
               Index           =   24
               Left            =   5535
               TabIndex        =   58
               Top             =   1170
               Width           =   105
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
               Left            =   60
               TabIndex        =   57
               Top             =   1455
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Order Tc/Yy/Chr"
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
               Left            =   60
               TabIndex        =   56
               Top             =   1170
               Width           =   1785
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Curr Location"
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
               TabIndex        =   55
               Top             =   2880
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
               Left            =   6840
               TabIndex        =   54
               Top             =   480
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
               Left            =   4395
               TabIndex        =   53
               Top             =   480
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
               Left            =   1845
               TabIndex        =   52
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBagHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Report --------'
' This repport is used to show the history of a bag. i.e. it shows the transactions of a bag
' and the details of the raw materials added or subtracted from the bag in a chronological
' fashion
'--------------------

Option Explicit
Dim Rep As New EmrRepBagHist
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mb_CTBMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
Dim ms_BHDets As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  'Keep This Section
End Sub

Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)
'*** Jay 2.13(CT) *** (BOpnPtQty, BRjPtQty added)
' ***** Manali 3.03 - qRmSzDesc included
  '*** Report Sql ***
'  Select BYy, BChr, BNo, BOpnQty, BOdTc, BOdYy, BOdChr, BOdNo, BOdSr, BOpnLoc,
'  BOpnDt, BLoc, BWrk, BQty, BRjQty, BOpnPtQty, BRjPtQty, BPtQty, BGrWt, 0 as NetWt, space(8) as LstRecDt, BCls,
'  OmCmCd, OdOmCtCd, OdPrdSeq, space(225) as qSeq, TdRmCd as qMainCd, BOdDmCd, BOdSfx, BOdDmSz,
'  OdDmCol, space(23) as qInvNo, space(23) as qCTBNo,
'  TDt, TdTc, TdYy, TdChr, TdNo, TdSr, TdSrNo, TdYy as qTdYy, TdKey, FdYy as qFdYy, FdKey, FmdYy as qFmdYy, FmdKey,
'  Txnd.ModUsr as qUsrCd, Txnd.ModTime as qModTime, TdFrBLoc, TdWrk, TdToBLoc, TdByWrk, RmCtg, TdRmCd as qRmCd,
'  TdRmSz as qRmSz, RisRmSzDesc as qRmSzDesc, TdRmStkRt as qStkRt, PValue3 as qPtrYN, TdRmQty as qRmQty,
'  TdRmWt as qRmWt, 0 as qIn, 0 as qOut, 0.0 as qPcsIn, 0 as qPrtIn, 0.0 as qPcsOut,
'  0 as qPrtOut, 0.0 as qPcsRej, 0.0 as qPcsPtRej, 0.0 as qIssQty, 0 as qIssPtQty,
'  0.0 as qIssWt, TdRmCd as qLossRmCd, 0.0 as qLossRmWt, TdBYy As qOthBYy,
'  TdBChr As qOthBChr, TdBNo As qOthBNo, space(4) As qFlag, BFgSubLoc
'  from Txnd, Bag, OrdDsg, OrdMst, Txn, Param, RmMst, Fgd where 1= 2
  '*** Report Sql ***
'*** Jay 2.13(CT) ***

  '*** Sub Report Sql ***
  'Select BYy, BChr, BNo, RmCtg, RmSCtg, TdRmCd, TdRmQty, TdRmWt, 0 as qSrt
  'from Bag, Txnd, RmMst where 1= 2
  '*** Sub Report Sql ***
   
  Dim wrepcnd As String, wCnd As String, wCommSel As String, wSel1 As String, wLnk1 As String
  Dim wSel2 As String, wLnk2 As String, wSel3 As String, wLnk3 As String, wSqlStrg As String
  Dim wCndFg As String, wSel4 As String, wLnk4 As String, wCtCd As String, wSel5 As String, wLnk5 As String
  Dim wSel6 As String, wLnk6 As String, wSel7 As String, wLnk7 As String  'for repair
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  
  '****** Sachin 3.02.0 - New fields Condition
  Dim wAddlCnd As String
  
  '*** Jay 2.13(CT) ***
  ' **** Zubin 211 **** '
  Dim ws_BagJoin As String
  ws_BagJoin = " Join Param BCHR On BCHR.PTyp= 'BCHR' and BCHR.PMCd= BChr and BCHR.PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and BCHR.PValue= '" + adc("UoBagTyp") + "' ", " ")
  'If UCase(adc("UoBagTyp")) = "N" Then
  '  ws_BagJoin = " Join Param BCHR On BCHR.PTyp= 'BCHR' and BCHR.PMCd= BChr and BCHR.PSCd= '' and BCHR.PValue = 'N' "
  '            'IIF(adc("UOBCHRFR") <> "", " and BCHR.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
  'ElseIf UCase(adc("UoBagTyp")) = "Y" Then
  '  ws_BagJoin = " Join Param BCHR On BCHR.PTyp= 'BCHR' and BCHR.PMCd= BChr and BCHR.PSCd= '' and BCHR.PValue = 'Y' "
  '            'IIF(adc("UOBCHRFR") <> "", " and BCHR.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
  'Else
  '  ws_BagJoin = " Join Param BCHR On BCHR.PTyp= 'BCHR' and BCHR.PMCd= BChr and BCHR.PSCd= '' "
  '            'IIF(adc("UOBCHRFR") <> "", " and BCHR.PMCd= '" + adc("UOBCHRFR") + "' ", " ")
  'End If
  ' **** Zubin 211 **** '
  '*** Jay 2.13(CT) ***
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  'Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
    
  '### find out values for NetWt, LstRecDt, qFlag
  
  If UCase(adc.MenuCd) = UCase("RprHis") Then
    If adc("UOOMTCFR") <> "" Then adc("UOOMTCFR").CmpStr = "BRepODTc >="
    If adc("UOOMYYFR") <> "" Then adc("UOOMYYFR").CmpStr = "BRepOdYy >="
    If adc("UOOMCHRFR") <> "" Then adc("UOOMCHRFR").CmpStr = "BRepOdChr >="
    If adc("UOOMNOFR") <> 0 Then adc("UOOMNOFR").CmpStr = "BRepOdNo >="
    If adc("UOODSRFR") <> 0 Then adc("UOODSRFR").CmpStr = "BRepOdSr >="
    If adc("UOOMTCTO") <> "" Then adc("UOOMTCTO").CmpStr = "BRepODTc <="
    If adc("UOOMYYTO") <> "" Then adc("UOOMYYTO").CmpStr = "BRepOdYy <="
    If adc("UOOMCHRTO") <> "" Then adc("UOOMCHRTO").CmpStr = "BRepOdChr <="
    If adc("UOOMNOTO") <> 0 Then adc("UOOMNOTO").CmpStr = "BRepOdNo <="
    If adc("UOODSRTO") <> 0 Then adc("UOODSRTO").CmpStr = "BRepOdSr <="
    If adc("UOISSLOCFR") <> 0 Then adc("UOISSLOCFR").CmpStr = "BRepLoc >="
    If adc("UOISSLOCTO") <> 0 Then adc("UOISSLOCTO").CmpStr = "BRepLoc <="
    If adc("UOISSLOCSEL") <> 0 Then adc("UOISSLOCSEL").CmpStr = "BRepLoc In "
  Else
    If adc("UOOMTCFR") <> "" Then adc("UOOMTCFR").CmpStr = "BOdTc >="
    If adc("UOOMYYFR") <> "" Then adc("UOOMYYFR").CmpStr = "BOdYy >="
    If adc("UOOMCHRFR") <> "" Then adc("UOOMCHRFR").CmpStr = "BOdChr >="
    If adc("UOOMNOFR") <> 0 Then adc("UOOMNOFR").CmpStr = "BOdNo >="
    If adc("UOODSRFR") <> 0 Then adc("UOODSRFR").CmpStr = "BOdSr >="
    If adc("UOOMTCTO") <> "" Then adc("UOOMTCTO").CmpStr = "BOdTc <="
    If adc("UOOMYYTO") <> "" Then adc("UOOMYYTO").CmpStr = "BOdYy <="
    If adc("UOOMCHRTO") <> "" Then adc("UOOMCHRTO").CmpStr = "BOdChr <="
    If adc("UOOMNOTO") <> 0 Then adc("UOOMNOTO").CmpStr = "BOdNo <="
    If adc("UOODSRTO") <> 0 Then adc("UOODSRTO").CmpStr = "BOdSr <="
    If adc("UOISSLOCFR") <> 0 Then adc("UOISSLOCFR").CmpStr = "BLoc >="
    If adc("UOISSLOCTO") <> 0 Then adc("UOISSLOCTO").CmpStr = "BLoc <="
    If adc("UOISSLOCSEL") <> 0 Then adc("UOISSLOCSEL").CmpStr = "BLoc In "
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCndFg = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  
    
  '****** Sachin 3.02.0 - New Fields Condition
  If Trim(adc("UOCMCDFR")) <> "" Then wAddlCnd = " OdOmCmCd >= '" + adc("UOCMCDFR") + "'"
  If Trim(adc("UOCMCDTO")) <> "" Then wAddlCnd = IIF(wAddlCnd <> "", wAddlCnd + " And ", "") + " OdOmCmCd <= '" + adc("UOCMCDTO") + "'"
  If Trim(adc("UOCMCDSEL")) <> "" Then wAddlCnd = IIF(wAddlCnd <> "", wAddlCnd + " And ", "") + " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")"
  
  If mb_CTBMod = True Then
    '****** Sachin 3.02 [26/11/07] - Partition Fields And Id fields added below ******
    wCtCd = " IsNull((Select InCmCd From InvHd Join InvFgd On InIdNo=IfInIdNo and InCoCd=IfCoCd  " + _
             "and InTc=IfTc and  InYy=IfYy and InChr=IfChr and InNo=IfNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=InPrtKey ", "") + " Join Param  " + _
             "On PTyp = 'TC' And PmCd = IfTc And PValue = 'CTB' Where IfCoCd= BCoCd  " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + "), BOdCtCd) as OdOmCtCd, "
  Else
    wCtCd = " '' as OdOmCtCd, "
  End If
  
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  wRmSzDesc = "(case when (Select PDesc225 from Param where PTyp='RMSCTG' and PMCd=IsNull(RisRmCtg, '') and PSCd=RmSCtg)='Y' then IsNull(RisRmSzDesc, '') Else '' End)"
  ' **** Manali 3.03 - 08/07/08 - RmSzDesc
  '**Bhavna added picnm
  ' **** Zubin 211 **** 'added PartQty fields
  '****** Sachin 3.02 [26/11/07] - Partition Fields And Id fields added below ******
  ' RmaRam
  'wCommSel = " BYy, BChr, BNo,BOpnQty,brepODtc as BOdTc,brepYY as BOdYy,brepCHR as BOdChr,brepNO as BOdNo,brepSR as BOdSr, OdPicNm, " + _
             "BOpnLoc, BOpnDt, BLoc, BWrk, BQty, BRjQty, BOpnPtQty, BRjPtQty, BPtQty, BGrWt, '*' as NetWt, " + _
             "'*' as LstRecDt, BCls, OmCmCd, " & wCtCd & " OdPrdSeq, " + _
             "(Select vPDesc225 From vParam Where vPCoCd= OdCoCd and " + _
             "vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq) as qSeq, " + _
             "(Select OrRmCd from OrdRm where OrCoCd = OdCoCd and " + _
             "OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + _
             "and OrMainMet= 'Y' " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + " ) as qMainCd, BOdDmCd, BOdSfx, BOdDmSz, OdDmCol, " + _
             "IsNull((Select IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ LTrim(Str(IfNo))+ '/'+ LTrim(Str(IfSr)) " + _
             "from InvFgd Join Param On PTyp = 'TC' And PmCd = IfTc And PValue = 'IN' where IfCoCd= BCoCd " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + "  ), '') as qInvNo, " + _
             IIF(mb_CTBMod = True, "IsNull((Select IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ LTrim(Str(IfNo))+ '/'+ LTrim(Str(IfSr)) " + _
             "from InvFgd Join Param On PTyp = 'TC' And PmCd = IfTc And PValue = 'CTB' where IfCoCd= BCoCd " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + " ), '') as qCTBNo, ", "'' as qCTBNo, ")
                                            
  'Repair - field selection changed by menuCode
  'Repair Jri,Jro change-OdRepGrWt added for repair gross wt In
  'HKS.9-Tree number added in select list
  '6.1b BOpnDesc added
  wCommSel = " BYy, BChr, BNo," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "OdOrdQty", "BOpnQty") + " as BOpnQty," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepODTc", "BOdTc") + " as BOdTc," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepOdYy", "BOdYy") + " as BOdYy," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepOdChr", "BOdChr") + " as BOdChr," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepOdNo", "BOdNo") + " as BOdNo," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepOdSr", "BOdSr") + " as BOdSr, OdPicNm, " + _
             "BOpnLoc,BOpnDesc," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "case when OdTc='JRO' then (select od.OdOmDt from OrdDsg od where od.OdTc ='JRI' and od.OdIdNo =OrdDsg.OdBaseIdKey) else OMDt end", "BOpnDt") + " as BOpnDt, " + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepLoc", "BLoc") + " as BLoc, BWrk," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "OdOrdQty", "BQty") + " as BQty, BRjQty," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "0", "BOpnPtQty") + " as BOpnPtQty, BRjPtQty, BPtQty, " + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "BRepGrWt", "BGrWt") + " as BGrWt, '*' as NetWt, " + _
             "'*' as LstRecDt," + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "case when BRepYN='Y' then 'N' else 'Y' end ", "BCls") + " as BCls, OmCmCd, " & wCtCd & " OdPrdSeq,Isnull(OdRepGrWt,0) as OdRepGrWt,  " + _
             "(Select vPDesc225 From vParam Where vPCoCd= OdCoCd and " + _
             "vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq) as qSeq, " + _
             "(Select OrRmCd from OrdRm where OrCoCd = OdCoCd and " + _
             "OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr " + _
             "and OrMainMet= 'Y' " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + " ) as qMainCd, BOdDmCd, BOdSfx, BOdDmSz, OdDmCol," + _
             IIF(UCase(adc.MenuCd) = UCase("RprHis"), "(Select OdTc+ '/'+ OdYy+ '/'+ OdChr+ '/'+ LTrim(Str(OdNo))+ '/'+ LTrim(Str(OdSr)) from OrdDsg ou where ou.OdBaseIdKey =orddsg.OdIdNo)", _
             "IsNull((Select IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ LTrim(Str(IfNo))+ '/'+ LTrim(Str(IfSr)) " + _
             "from InvFgd Join Param On PTyp = 'TC' And PmCd = IfTc And PValue = 'IN' where IfCoCd= BCoCd " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + " ), '')") + " as qInvNo," + _
             IIF(mb_CTBMod = True, "IsNull((Select IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ LTrim(Str(IfNo))+ '/'+ LTrim(Str(IfSr)) " + _
             "from InvFgd Join Param On PTyp = 'TC' And PmCd = IfTc And PValue = 'CTB' where IfCoCd= BCoCd " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + " ), '') as qCTBNo, ", "'' as qCTBNo, ") + _
             "isnull((select top 1 TrYmd+'/'+TrChr+'/'+LTrim(str(TrNo)) from TreeMst Join TreeDet on TrCoCd=TrdCoCd and TrYmd=TrdYmd and TrChr=TrdChr " + _
             "and TrNo=TrdNo where BCoCd=TrdCoCd and BYy=TrdByy and BChr=TrdBChr and BNo=TrdBNo order by TrDt,TrdSr),'') TreeNo, "

  'uma
  '             '"(Select ou.OdTc+'/'+ou.OdYy+'/'+ou.OdChr+'/'+LTrim(Str(ou.OdNo)) from orddsg ou join bag on byy=ou.OdDmStkYy and bchr=ou.OdDmStkChr and bno=ou.OdDmStkNo where ou.odtc='JRO' )  as qInvNo, "
  '"IsNull((Select IfTc+ '/'+ IfYy+ '/'+ IfChr+ '/'+ LTrim(Str(IfNo))+ '/'+ LTrim(Str(IfSr)) " + _
             "from InvFgd Join Param On PTyp = 'TC' And PmCd = IfTc And PValue = 'IN' where IfCoCd= BCoCd " + _
             "and IfFdBYy= BYy and IfFdBChr= BChr and IfFdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And IfPrtKey=BPrtKey ", "") + "  ), '') as qInvNo, " + _

  'case when OdTc='JRO' then (select od.OdOmDt from OrdDsg od where od.OdTc ='JRI' and od.OdIdNo =OrdDsg.OdBaseIdKey) else OMDt end"
  'Repair - Join OrdDsg for normal bag
  Dim ws_OdBagJoin As String
  ws_OdBagJoin = "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
          "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "")
  
  'Repair - Join OrdDsg for repair bag
  If UCase(adc.MenuCd) = UCase("RprHis") Then _
   ws_OdBagJoin = "Join OrdDsg on OdCoCd= BCoCd and OdTc = BRepOdTc And OdYy = BRepOdYy And OdChr= BRepOdChr And OdNo = BRepOdNo and OdSr= BRepOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "")

'  ws_OdBagJoin = "Join OrdDsg on OdCoCd= BCoCd and OdDmStkYy = BYy And OdDmStkChr = BCHR And OdDmStkNo = BNo and OdTc='JRI'" + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "")
  '6.1b-41 if complete history is selected then no need to check Repair chr YN
  Dim ws_vParamJoin As String
  ws_vParamJoin = " join vParam on vPTyp ='CHR' and vPCoCd =Pt.TdCoCd and vPMCd =Pt.TdTc and vPSCd =Pt.TdChr " + _
                  IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "and vPValue3='" + _
                  IIF(UCase(adc.MenuCd) = UCase("RprHis"), "Y", "N") + "'")

  ' **** Zubin 211 **** 'added PartQty fields

  ' Manoj 2.10.0
  '***Bhavna added tdlotno with tdrmcd in all sql
  '*** Sql for all transactions other than TcTyp= 'BS' ***
  '(Select RmCtg from RmMst where RmCd= Ch.TdRmCd)
  ' **** Zubin 211 changed **** '
  ' *** Manali 3.03 - 08/07/08 - qRmSzDesc added below
  wSel1 = " TDt, Pt.TdTc, Pt.TdYy, Pt.TdChr, Pt.TdNo, Pt.TdSr, IsNull(Ch.TdSrNo, 0) as TdSrNo, " + _
          "Pt.TdYy as qTdYy, Pt.TdKey, '' as qFdYy, 0 as FdKey, '' as qFmdYy, 0 as FmdKey, Pt.ModUsr as qUsrCd, Pt.ModTime As qModTime, Pt.TdFrBLoc, " + _
          "(case when Prm.PValue= 'BV' then '' else Pt.TdWrk end) as TdWrk, Pt.TdToBLoc, " + _
          "(case when Prm.PValue= 'BV' then Pt.TdWrk else Pt.TdByWrk end) as TdByWrk, " + _
          "IsNull(RmCtg, '') as RmCtg, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmCd else '' end), '') as qRmCd, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdLotNo else '' end), '') as qLotNo, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmSz else 0 end), 0) as qRmSz, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmSz2 else 0 end), 0) as qRmSz2, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmSz3 else 0 end), 0) as qRmSz3, " + _
          wRmSzDesc + " as qRmSzDesc, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmStkRt else 0 end), 0) as qStkRt, " + _
          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
          "IsNull((case when Prm.PValue<> 'BV' then Ch.TdRmQty else 0 end), 0) as qRmQty, " + _
          "IsNull((case when Prm.PValue<> 'BV' and not exists(Select LocTyp from Loc where " + _
          "LocCoCd= Ch.TdCoCd and LocCd in (Ch.TdFrRmLoc, Ch.TdToRmLoc) and LocTyp= 'L') then Ch.TdRmWt else 0 end), 0) as qRmWt, " + _
          "IsNull((case when Prm.PValue<> 'BV' and (((Ch.TdFrRmLoc= 'W' OR Ch.TdFrRmLoc= 'XR') and Ch.TdFrRmDc= 'D') or ((Ch.TdToRmLoc= 'W' OR Ch.TdToRmLoc= 'XR') and Ch.TdToRmDc= 'D')) then 1 else 0 end), 0) as qIn, " + _
          "IsNull((case when Prm.PValue<> 'BV' and (((Ch.TdFrRmLoc= 'W' OR Ch.TdFrRmLoc= 'XR') and Ch.TdFrRmDc= 'C') or ((Ch.TdToRmLoc= 'W' OR Ch.TdToRmLoc= 'XR') and Ch.TdToRmDc= 'C')) then 1 else 0 end), 0) as qOut, "

          '"IsNull((case when Prm.PValue<> 'BV' and ((Ch.TdFrRmLoc= 'W' and Ch.TdFrRmDc= 'D') or (Ch.TdToRmLoc= 'W' and Ch.TdToRmDc= 'D')) then 1 else 0 end), 0) as qIn, " + _
          "IsNull((case when Prm.PValue<> 'BV' and ((Ch.TdFrRmLoc= 'W' and Ch.TdFrRmDc= 'C') or (Ch.TdToRmLoc= 'W' and Ch.TdToRmDc= 'C')) then 1 else 0 end), 0) as qOut, "

  '****** Sachin 3.02 [26/11/07] - Partition Fields And Id fields added below ******
  wSel1 = wSel1 + "0 as qPcsIn, 0 as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, Pt.TdRjQty as qPcsRej, " + _
          "Pt.TdRjPtQty as qPcsPtRej, " + _
          "(case when Prm.PValue= 'BV' then Pt.TdBQty else '' end) as qIssQty, " + _
          "(case when Prm.PValue= 'BV' then Pt.TdPtQty else '' end) as qIssPtQty, " + _
          "(case when Prm.PValue= 'BV' then Pt.TdBGrWt else '' end) as qIssWt, " + _
          "IsNull((case when (Select LocTyp from Loc where LocCoCd= Ch.TdCoCd and LocCd= Ch.TdToRmLoc)= 'L' " + _
          "then Ch.TdRmCd when (Select LocTyp from Loc where LocCoCd= Ch.TdCoCd and LocCd= Ch.TdFrRmLoc)= 'L' " + _
          "then Ch.TdRmCd else '' end) , '') as qLossRmCd, " + _
          "IsNull((case when (Select LocTyp from Loc where LocCoCd= Ch.TdCoCd and LocCd= Ch.TdToRmLoc)= 'L' " + _
          "then Ch.TdRmWt*(case when Ch.TdToRmDc= 'D' then 1 else -1 end) " + _
          "when (Select LocTyp from Loc where LocCoCd= Ch.TdCoCd and LocCd= Ch.TdFrRmLoc)= 'L' " + _
          "then Ch.TdRmWt*(case when Ch.TdFrRmDc= 'D' then 1 else -1 end) else 0 end) , 0) as qLossRmWt, " + _
          "'' as qOthBYy, '' as qOthBChr, 0 as qOthBNo, '*' as qFlag," + IIF(UCase(adc.MenuCd) <> UCase("RprHis"), "BFgSubLoc", "''") + " as  BFgSubLoc "
  'wLnk1 = " from Txnd Pt Join Bag on BIdNo=Pt.TdBIdNo and BCoCd= Pt.TdCoCd and BYy= Pt.TdBYy and BChr= Pt.TdBChr and BNo= Pt.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Pt.TdPrtKey ", "") + _
          ws_BagJoin + _
          "Join OrdDsg on OdIdNo=BOdIdNo and OdCoCd= BCoCd and OdTc= BOdTc and OdYy= BOdYy and " + _
          "OdChr= BOdChr and OdNo= BOdNo and OdSr= BOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "") + _
          "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
          "Join Txn on TIdNo=Pt.TdTIdNo and TCoCd= Pt.TdCoCd and TTc= Pt.TdTc and TYy= Pt.TdYy and TChr= Pt.TdChr and TNo= Pt.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Pt.TdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= Pt.TdTc " + _
          "join vParam on vPTyp ='CHR' and vPCoCd =Pt.TdCoCd and vPMCd =Pt.TdTc and vPSCd =Pt.TdChr and vPValue3='" + IIF(UCase(adc.MenuCd) = UCase("RprHis"), "N", "Y") + "'" + _
          "Left outer join Txnd Ch on Ch.TdCoCd= Pt.TdCoCd and Ch.TdTc= Pt.TdTc and Ch.TdYy= Pt.TdYy and " + _
          "Ch.TdChr= Pt.TdChr and Ch.TdNo= Pt.TdNo and Ch.TdSr= Pt.TdSr and Ch.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey=Ch.TdPrtKey ", "") + _
          "Left outer join RmMst on RmCd= Ch.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=Ch.TdPrtKey ", "") + _
          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=Ch.TdRmSz " + _
          "where Pt.TdSrNo=0 and Prm.PValue<> 'BS' " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey='" + ctCurrPrtn + "'", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wCnd
  '6.1b-41 ws_vParamJoin added and if full history is selected then date check in txn is removed
  wLnk1 = " from Txnd Pt Join Bag on BIdNo=Pt.TdBIdNo and BCoCd= Pt.TdCoCd and BYy= Pt.TdBYy and BChr= Pt.TdBChr and BNo= Pt.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Pt.TdPrtKey ", "") + _
          ws_BagJoin + ws_OdBagJoin + _
          "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
          "Join Txn on " + IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "TDt >=OmDt  and ") + " TIdNo=Pt.TdTIdNo and TCoCd= Pt.TdCoCd and TTc= Pt.TdTc and TYy= Pt.TdYy and TChr= Pt.TdChr and TNo= Pt.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Pt.TdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= Pt.TdTc " + _
          ws_vParamJoin + _
          "Left outer join Txnd Ch on Ch.TdCoCd= Pt.TdCoCd and Ch.TdTc= Pt.TdTc and Ch.TdYy= Pt.TdYy and " + _
          "Ch.TdChr= Pt.TdChr and Ch.TdNo= Pt.TdNo and Ch.TdSr= Pt.TdSr and Ch.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey=Ch.TdPrtKey ", "") + _
          "Left outer join RmMst on RmCd= Ch.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=Ch.TdPrtKey ", "") + _
          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=Ch.TdRmSz " + _
          "where Pt.TdSrNo=0 and Prm.PValue<> 'BS' " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey='" + ctCurrPrtn + "'", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wCnd
      
  
  ' **** Zubin 211 changed **** '
  
          '"IsNull((case when PValue<> 'BV' then Ch.TdRmWt else 0 end), 0) as qRmWt, " + _
          '"IsNull((case when PValue= 'BV' then Ch.TdRmCd else '' end), '') as qLossRmCd, " + _
          '"IsNull((case when PValue= 'BV' then Ch.TdRmWt else '' end), '') as qLossRmWt, " + _

  '*** Sql for all transactions of TcTyp= 'BS' where Pcs go out of Bag ***
  ' **** Zubin 211 changed **** '
  ' *** Manali 3.03 - 08/07/08 - qRmSzDesc added below
  'SriTmp

  wSel2 = " TDt, Pt.TdTc, Pt.TdYy, Pt.TdChr, Pt.TdNo, IsNull(OthPt.TdSr, 0) as TdSr, " + _
          "IsNull(OthCh.TdSrNo, 0) as TdSrNo, Pt.TdYy as qTdYy, Pt.TdKey, '' as qFdYy, 0 as FdKey, '' as qFmdYy, 0 as FmdKey, " + _
          "IsNull(OthPt.ModUsr, '') as qUsrCd, IsNull(OthPt.ModTime, 0) as qModTime, Pt.TdFrBLoc, Pt.TdWrk, Pt.TdToBLoc, Pt.TdByWrk, " + _
          "IsNull(RmCtg, '') as RmCtg, IsNull(OthCh.TdRmCd, '') as qRmCd, IsNull(OthCh.TdLotNo, '') as qLotNo, " + _
          "IsNull(OthCh.TdRmSz, 0) as qRmSz, " + _
          "IsNull(OthCh.TdRmSz2, 0) as qRmSz2, " + _
          "IsNull(OthCh.TdRmSz3, 0) as qRmSz3, " + _
          wRmSzDesc + " as qRmSzDesc, " + _
          "IsNull(OthCh.TdRmStkRt, 0) as qStkRt, " + _
          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
          "IsNull(OthCh.TdRmQty, 0) as qRmQty, IsNull(OthCh.TdRmWt, 0) as qRmWt, " + _
          "(case when OthCh.TdFrRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdFrRmDc= 'C') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'D')) then 1 " + _
          "      when OthCh.TdToRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdToRmDc= 'C') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'D')) then 1 " + _
          " else 0 end) as qIn, " + _
          "(case when OthCh.TdFrRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdFrRmDc= 'D') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'C')) then 1 " + _
          "      when OthCh.TdToRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdToRmDc= 'D') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'C')) then 1 else 0 end) as qOut, " + _
          "0 as qPcsIn, 0 as qPrtIn, IsNull(OthPt.TdBQty, 0) as qPcsOut, " + _
          "IsNull(OthPt.TdPtQty, 0) as qPrtOut, 0 as qPcsRej, 0 as qPcsPtRej, 0 as qIssQty, " + _
          "0 as qIssPtQty, 0 as qIssWt, '' as qLossRmCd, 0 as qLossRmWt, " + _
          "IsNull(OthPtForBagNo.TdBYy,isnull(OthPt.TdBYy,'')) as qOthBYy, IsNull(OthPt.TdBChr, '') as qOthBChr, " + _
          "IsNull(OthPtForBagNo.TdBNo,isnull(OthPt.TdBNo,'')) as qOthBNo, " + _
          "'*' as qFlag, BFgSubLoc "
  '****** Sachin 3.02 - Id fields in joins ******
  ' *** Manali 3.03
  '6.1b-41 ws_vParamJoin added and if full history is selected then date check in txn is removed
  wLnk2 = " from Txnd Pt " + _
          "Join Bag on BIdNo=Pt.TdBIdNo and BCoCd= Pt.TdCoCd and BYy= Pt.TdBYy and BChr= Pt.TdBChr and BNo= Pt.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Pt.TdPrtKey ", "") + _
          ws_BagJoin + ws_OdBagJoin + _
          "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
          "Join Txn on " + IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "TDt >=OmDt and ") + " TIdNo=Pt.TdTIdNo and TCoCd= Pt.TdCoCd and TTc= Pt.TdTc and TYy= Pt.TdYy and TChr= Pt.TdChr and TNo= Pt.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Pt.TdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= Pt.TdTc " + _
          ws_vParamJoin + _
          "left outer join Txnd OthPt on OthPt.TdCoCd= Pt.TdCoCd and OthPt.TdTc= Pt.TdTc " + _
          "and OthPt.TdYy= Pt.TdYy and OthPt.TdChr= Pt.TdChr and OthPt.TdNo= Pt.TdNo " + _
          "and OthPt.TdSrNo= 0 and ((Pt.TdRefKey=0 and Pt.TdRefYy=''  and OthPt.TdRefKey=Pt.TdKey and OthPt.TdRefYy=Pt.TdYy) or(Pt.TdRefKey<>0 and Pt.TdRefYy<>''   and Pt.TdKey=OthPt.TdKey and Pt.TdYy=OthPt.TdYy)) " + IIF(gs_Partition = ctCurrPrtn, " And OthPt.TdPrtKey=Pt.TdPrtKey ", "") + _
          "left outer join Txnd OthCh on OthCh.TdCoCd= OthPt.TdCoCd and OthCh.TdTc= OthPt.TdTc " + _
          "and OthCh.TdYy= OthPt.TdYy and OthCh.TdChr= OthPt.TdChr and OthCh.TdNo= OthPt.TdNo " + _
          "and OthCh.TdSr= OthPt.TdSr and OthCh.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And OthCh.TdPrtKey=OthPt.TdPrtKey ", "") + _
          "left outer join Txnd OthPtForBagNo on OthPtForBagNo.TdCoCd= Pt.TdCoCd and OthPtForBagNo.TdTc= Pt.TdTc " + _
          "and OthPtForBagNo.TdYy= Pt.TdYy and OthPtForBagNo.TdChr= Pt.TdChr and OthPtForBagNo.TdNo= Pt.TdNo " + _
          "and OthPtForBagNo.TdSrNo= 0 and ((Pt.TdRefKey<>0 and Pt.TdRefYy<>''  and OthPtForBagNo.TdKey=Pt.TdRefKey and OthPtForBagNo.TdYy=Pt.TdRefYy)) " + IIF(gs_Partition = ctCurrPrtn, " And OthPt.TdPrtKey=Pt.TdPrtKey ", "") + _
          "left outer join RmMst on RmCd= OthCh.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OthCh.TdPrtKey ", "") + _
          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OthCh.TdRmSz " + _
          "where Pt.TdSrNo=0 and Prm.PValue= 'BS' " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey='" + ctCurrPrtn + "' ", "") + " And " + _
          "Pt.TdBDc= 'C' " + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wCnd
  ' **** Zubin 211 changed **** '
      
  '*** Sql for all transactions of TcTyp= 'BS' where Pcs come into Bag ***
  ' **** Zubin 211 changed **** '
  ' *** Manali 3.03 - 08/07/08 - qRmSzDesc added below
'  wSel3 = " TDt, Pt.TdTc, Pt.TdYy, Pt.TdChr, Pt.TdNo, Pt.TdSr, IsNull(Ch.TdSrNo, 0) as TdSrNo, " + _
          "Pt.TdYy as qTdYy, Pt.TdKey, '' as qFdYy, 0 as FdKey, '' as qFmdYy, 0 as FmdKey, " + _
          "Pt.ModUsr as qUsrCd, Pt.ModTime as qModTime, Pt.TdFrBLoc, Pt.TdWrk, Pt.TdToBLoc, Pt.TdByWrk, " + _
          "IsNull(RmCtg, '') as RmCtg, IsNull(Ch.TdRmCd, '') as qRmCd,  IsNull(Ch.TdLotNo, '') as qLotNo, " + _
          "IsNull(Ch.TdRmSz, 0) as qRmSz, " + _
          "IsNull(Ch.TdRmSz2, 0) as qRmSz2, " + _
          "IsNull(Ch.TdRmSz3, 0) as qRmSz3, " + _
          wRmSzDesc + " as qRmSzDesc, " + _
          "IsNull(Ch.TdRmStkRt, 0) as qStkRt, " + _
          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
          "IsNull(Ch.TdRmQty, 0) as qRmQty, IsNull(Ch.TdRmWt, 0) as qRmWt, " + _
          "(case when Ch.TdFrRmLoc= 'W' and Ch.TdFrRmDc= 'D' then 1 " + _
          "      when Ch.TdToRmLoc= 'W' and Ch.TdToRmDc= 'D' then 1 else 0 end) as qIn, " + _
          "(case when Ch.TdFrRmLoc= 'W' and Ch.TdFrRmDc= 'C' then 1 " + _
          "      when Ch.TdToRmLoc= 'W' and Ch.TdToRmDc= 'C' then 1 else 0 end) as qOut, " + _
          "Pt.TdBQty as qPcsIn, Pt.TdPtQty as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, " + _
          "0 as qPcsRej, 0 as qPcsPtRej, 0 as qIssQty, 0 as qIssPtQty, 0 as qIssWt, " + _
          "'' as qLossRmCd, 0 as qLossRmWt, IsNull(OthPt.TdBYy, '') as qOthBYy, IsNull(OthPt.TdBChr, '') as qOthBChr, " + _
          "IsNull(OthPt.TdBNo, 0) as qOthBNo, '*' as qFlag, BFgSubLoc "
  '****** Sachin 3.02 - Id fields in joins ******
  wSel3 = " TDt, Pt.TdTc, Pt.TdYy, Pt.TdChr, Pt.TdNo, IsNull(OthPt.TdSr, 0) as TdSr, " + _
          "IsNull(OthCh.TdSrNo, 0) as TdSrNo, Pt.TdYy as qTdYy, Pt.TdKey, '' as qFdYy, 0 as FdKey, '' as qFmdYy, 0 as FmdKey, " + _
          "IsNull(OthPt.ModUsr, '') as qUsrCd, IsNull(OthPt.ModTime, 0) as qModTime, Pt.TdFrBLoc, Pt.TdWrk, Pt.TdToBLoc, Pt.TdByWrk, " + _
          "IsNull(RmCtg, '') as RmCtg, IsNull(OthCh.TdRmCd, '') as qRmCd, IsNull(OthCh.TdLotNo, '') as qLotNo, " + _
          "IsNull(OthCh.TdRmSz, 0) as qRmSz, " + _
          "IsNull(OthCh.TdRmSz2, 0) as qRmSz2, " + _
          "IsNull(OthCh.TdRmSz3, 0) as qRmSz3, " + _
          wRmSzDesc + " as qRmSzDesc, " + _
          "IsNull(OthCh.TdRmStkRt, 0) as qStkRt, " + _
          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
          "IsNull(OthCh.TdRmQty, 0) as qRmQty, IsNull(OthCh.TdRmWt, 0) as qRmWt, " + _
          "(case when OthCh.TdFrRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdFrRmDc= 'C') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'D')) then 1 " + _
          "      when OthCh.TdToRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdToRmDc= 'C') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'D')) then 1 " + _
          " else 0 end) as qIn, " + _
          "(case when OthCh.TdFrRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdFrRmDc= 'D') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'C')) then 1 " + _
          "      when OthCh.TdToRmLoc= 'W' and ((Pt.TdRefKey=0 and OthCh.TdToRmDc= 'D') or (Pt.TdRefKey<>0 and OthCh.TdFrRmDc= 'C')) then 1 else 0 end) as qOut, " + _
          "OthPt.TdBQty  as qPcsIn, OthPt.TdPtQty as qPrtIn, 0 as qPcsOut, " + _
          "0 as qPrtOut, 0 as qPcsRej, 0 as qPcsPtRej, 0 as qIssQty, " + _
          "0 as qIssPtQty, 0 as qIssWt, '' as qLossRmCd, 0 as qLossRmWt, " + _
          "IsNull(OthPtForBagNo.TdBYy,isnull(OthPt.TdBYy,'')) as qOthBYy, IsNull(OthPt.TdBChr, '') as qOthBChr, " + _
          "IsNull(OthPtForBagNo.TdBNo,isnull(OthPt.TdBNo,'')) as qOthBNo, " + _
          "'*' as qFlag, BFgSubLoc "
  '6.1b-41 ws_vParamJoin added and if full history is selected then date check in txn is removed
  wLnk3 = "from Txnd Pt " + _
          "Join Bag on BIdNo=Pt.TdBIdNo and BCoCd= Pt.TdCoCd and BYy= Pt.TdBYy and BChr= Pt.TdBChr and BNo= Pt.TdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=Pt.TdPrtKey ", "") + _
          ws_BagJoin + ws_OdBagJoin + _
          "Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
          "Join Txn on " + IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "TDt >=OmDt and ") + " TIdNo=Pt.TdTIdNo and TCoCd= Pt.TdCoCd and TTc= Pt.TdTc and TYy= Pt.TdYy and " + _
          "TChr= Pt.TdChr and TNo= Pt.TdNo " + IIF(gs_Partition = ctCurrPrtn, " And TPrtKey=Pt.TdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= Pt.TdTc " + _
          ws_vParamJoin + _
          "left outer join Txnd OthPt on OthPt.TdCoCd= Pt.TdCoCd and OthPt.TdTc= Pt.TdTc " + _
          "and OthPt.TdYy= Pt.TdYy and OthPt.TdChr= Pt.TdChr and OthPt.TdNo= Pt.TdNo " + _
          "and OthPt.TdSrNo= 0 and ((Pt.TdRefKey=0 and Pt.TdRefYy=''  and OthPt.TdRefKey=Pt.TdKey and OthPt.TdRefYy=Pt.TdYy) or (Pt.TdRefKey<>0 and Pt.TdRefYy<>''   and Pt.TdKey=OthPt.TdKey and Pt.TdYy=OthPt.TdYy)) " + IIF(gs_Partition = ctCurrPrtn, " And OthPt.TdPrtKey=Pt.TdPrtKey ", "") + _
          "left outer join Txnd OthCh on OthCh.TdCoCd= OthPt.TdCoCd and OthCh.TdTc= OthPt.TdTc " + _
          "and OthCh.TdYy= OthPt.TdYy and OthCh.TdChr= OthPt.TdChr and OthCh.TdNo= OthPt.TdNo " + _
          "and OthCh.TdSr= OthPt.TdSr and OthCh.TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And OthCh.TdPrtKey=OthPt.TdPrtKey ", "") + _
          "left outer join Txnd OthPtForBagNo on OthPtForBagNo.TdCoCd= Pt.TdCoCd and OthPtForBagNo.TdTc= Pt.TdTc " + _
          "and OthPtForBagNo.TdYy= Pt.TdYy and OthPtForBagNo.TdChr= Pt.TdChr and OthPtForBagNo.TdNo= Pt.TdNo " + _
          "and OthPtForBagNo.TdSrNo= 0 and ((Pt.TdRefKey<>0 and Pt.TdRefYy<>''  and OthPtForBagNo.TdKey=Pt.TdRefKey and OthPtForBagNo.TdYy=Pt.TdRefYy)) " + IIF(gs_Partition = ctCurrPrtn, " And OthPt.TdPrtKey=Pt.TdPrtKey ", "") + _
          "  left outer join RmMst on RmCd= OthCh.TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OthCh.TdPrtKey ", "") + _
          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OthCh.TdRmSz " + _
          "where Pt.TdSrNo=0 and Prm.PValue= 'BS' and Pt.TdBDc= 'D' " + IIF(gs_Partition = ctCurrPrtn, " And Pt.TdPrtKey='" + ctCurrPrtn + "' ", "") + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wCnd
  ' **** Zubin 211 changed **** '
    
  ' **** Zubin 211 (ws_BagJoin, Prm.Value prefix added below) **** '
  ' **** Sql for all Fg transactions **** '
  
  '*** (Bef 2.12)
  '(case when Prm.PValue= 'FB' then " + _
          "(case when IsNull(TdToBLoc, '')<> '' then IsNull(TdToBLoc, '') else " + _
          "(case when IsNull(TdFrBLoc, '')<> '' then IsNull(TdFrBLoc, '') else BOpnLoc end) end) " + _
          "when Prm.PValue= 'FWB' then FgFrBLoc else '' end) as TdFrBLoc
  '*** (Bef 2.12)
   
  '*** (Jen 2.12) Changed the Logic for 'From Bag Loc' ***
  '****** Sachin 3.02 - Id fields in joins ******
  Dim ws_FrBLoc As String
  ws_FrBLoc = "(case when Prm.PValue= 'FB' then " + _
              " Isnull((select F.FgToBLoc from Fgd Fwb " + _
              "         join Fg F on F.FgIdNo=Fwb.FdFgIdNo and F.FgCoCd= Fwb.FdCoCd and F.FgTc= Fwb.FdTc " + _
              "               and F.FgYy= Fwb.FdYy and F.FgChr= Fwb.FdChr " + _
              "               and F.FgNo= Fwb.FdNo " + IIF(gs_Partition = ctCurrPrtn, " And F.FgPrtKey=Fwb.FdPrtKey ", "") + _
              "         Where FWB.FdBIdNo=BIdNo and FWB.FdCoCd = BCoCd And FWB.FdBYy = BYy and Fwb.FdBChr= BChr " + _
              "         and Fwb.FdBNo= BNo and Fwb.FdBYyKey= Fd.FdBYyKey " + IIF(gs_Partition = ctCurrPrtn, " And Fwb.FdPrtKey=Fd.FdPrtKey ", "") + _
              "         and Fwb.FdYy+ Str(Fwb.FdKey)= " + _
              "             (Select max(Fwb1.FdYy+ Str(Fwb1.FdKey)) From Fgd Fwb1 " + _
              "             where Fwb1.FdTc in (Select PMCd from Param where PTyp= 'TC' and PValue= 'FWB') " + _
              "             and Fwb1.FdCoCd= BCoCd and Fwb1.FdBYy= BYy and Fwb1.FdBChr= BChr " + _
              "             and Fwb1.FdBNo= BNo " + IIF(gs_Partition = ctCurrPrtn, " And Fwb1.FdPrtKey=BPrtKey ", "") + _
              "             and (Fwb1.FdYy< Fd.FdYy or (Fwb1.FdYy= Fd.FdYy and Fwb1.FdKey< Fd.FdKey)) " + _
              "             )" + _
              "         ), " + _
              "(case when IsNull(TdToBLoc, '')<> '' then IsNull(TdToBLoc, '') else " + _
              "(case when IsNull(TdFrBLoc, '')<> '' then IsNull(TdFrBLoc, '') else BOpnLoc end) end)) " + _
              "when Prm.PValue= 'FWB' then FgFrBLoc else '' end) as TdFrBLoc "
  '*** (Jen 2.12) Changed the Logic for 'From Bag Loc' ***
  
  ' *** Manali 3.03 - 08/07/08 - qRmSzDesc added below
  '578-34 FdBYyKey, 3, 7 changed to FdBYyKey, 3, 8
  wSel4 = " FgDt as TDt, FdTc as TdTc, FdYy as TdYy, FdChr as TdChr, FdNo as TdNo, FdSr as TdSr, 0 as TdSrNo, " + _
          "SubString(FdBYyKey, 1, 2) as qTdYy, " + _
          "convert(int, SubString(FdBYyKey, 3, 8)) as TdKey, FdYy as qFdYy, FdKey, '' as qFmdYy, 0 as FmdKey, " + _
          "Fd.ModUsr as qUsrCd, Fd.ModTime as qModTime, " + ws_FrBLoc + ", '' as TdWrk, FgToBLoc as TdToBLoc, " + _
          "TdWrk as TdByWrk, '' as RmCtg, '' as qRmCd, '' as qLotNo, 0.0000 as qRmSz,  0.0000 as qRmSz2,  0.0000 as qRmSz3, '' as qRmSzDesc, 0 as qStkRt, '' as qPtrYN, " + _
          "0 as qRmQty, 0 as qRmWt, " + _
          "0 as qIn, 0 as qOut, 0 as qPcsIn, 0 as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, " + _
          "0 as qPcsRej, 0 as qPcsPtRej, FdQty as qIssQty, 0 as qIssPtQty, FdGrWt as qIssWt, " + _
          "'' as qLossRmCd, 0 as qLossRmWt, '' as qOthBYy, '' as qOthBChr, 0 as qOthBNo, '*' as qFlag, BFgSubLoc "
  '****** Sachin 3.02 - Id fields in joins ******
  '6.1b-41 ws_vParamJoin added and if full history is selected then date check in fg is removed
  wLnk4 = " from Fgd Fd Join Bag on BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=FdPrtKey ", "") + ws_BagJoin + _
          ws_OdBagJoin + _
          "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
          "Join Fg on " + IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "FgDt >=OmDt and ") + " FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " And FgPrtKey=FdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= FdTc " + _
          "join vParam on vPTyp ='CHR' and vPCoCd =FdCoCd and vPMCd =FdTc and vPSCd =FdChr " + _
          IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", " and vPValue3='" + _
          IIF(UCase(adc.MenuCd) = UCase("RprHis"), "Y", "N") + "'") + _
          "Left Outer Join Txnd on TdCoCd= FdCoCd and TdBYy=FdBYy And TdBChr=FdBChr And TdBNo=FdBNo and TdYyKey= FdBYyKey and TdSrNo= 0 " + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey=FdPrtKey ", "") + _
          wCndFg + IIF(wAddlCnd <> "", IIF(wCndFg <> "", " And ", " Where ") + wAddlCnd, "") + IIF(gs_Partition = ctCurrPrtn, " And FdPrtKey='" + ctCurrPrtn + "'", "")
  ' **** Zubin 211 (ws_BagJoin, Prm.Value prefix added above) **** '
          
  ''*** Sql for all Fg transactions with Rm Details ***
  ' *** Sachin 2.13.0 - 24/06/2006 ******
  ' *** Manali 3.03 - 08/07/08 - qRmSzDesc added below
  '578-34 FdBYyKey, 3, 7 changed to FdBYyKey, 3, 8
  wSel5 = " FmDt as TDt, FmdTc as TdTc, FmdYy as TdYy, FmdChr as TdChr, FmdNo as TdNo, FmdSr as TdSr, 0 as TdSrNo, " + _
          "SubString(FdBYyKey, 1, 2) as qTdYy, " + _
          "convert(int, SubString(FdBYyKey, 3, 8)) as TdKey, FdYy as qFdYy, FdKey, FmdYy as qFmdYy, FmdKey, " + _
          "Fd.ModUsr as qUsrCd, Fd.ModTime as qModTime, FmdFrFgSubLoc as TdFrBLoc, '' as TdWrk, FmdToFgSubLoc as TdToBLoc, " + _
          "'' as TdByWrk, '' as RmCtg, '' as qRmCd, '' as qLotNo, 0.0000 as qRmSz, 0.0000 as qRmSz2, 0.0000 as qRmSz3, '' as qRmSzDesc, 0 as qStkRt, '' as qPtrYN, " + _
          "0 as qRmQty, 0 as qRmWt, " + _
          "0 as qIn, 0 as qOut, 0 as qPcsIn, 0 as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, " + _
          "0 as qPcsRej, 0 as qPcsPtRej, FdQty as qIssQty, 0 as qIssPtQty, FdGrWt as qIssWt, " + _
          "'' as qLossRmCd, 0 as qLossRmWt, '' as qOthBYy, '' as qOthBChr, 0 as qOthBNo, '*' as qFlag, BFgSubLoc "
  '6.1b-41 ws_vParamJoin added and if full history is selected then date check in Fm is removed
  wLnk5 = " from Fmd Fd Join Bag on BCoCd= FmdCoCd and BYy= FmdBYy and BChr= FmdBChr and BNo= FmdBNo " + IIF(gs_Partition = ctCurrPrtn, " And BPrtKey=FmdPrtKey ", "") + ws_BagJoin + _
          ws_OdBagJoin + _
          "Join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
          "Join Fm on " + IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", "FmDt >=OmDt and ") + " FmIdNo=FmdFmIdNo and FmCoCd= FmdCoCd and FmTc= FmdTc and FmYy= FmdYy and FmChr= FmdChr and FmNo= FmdNo " + IIF(gs_Partition = ctCurrPrtn, " And FmPrtKey=FmdPrtKey ", "") + _
          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= FmdTc " + _
          "join vParam on vPTyp ='CHR' and vPCoCd =FmdCoCd and vPMCd =FmdTc and vPSCd =FmdChr " + _
          IIF(UCase(adc.MenuCd) = UCase("RprHis") And adc("UoYn2") = "Y", "", " and vPValue3='" + _
          IIF(UCase(adc.MenuCd) = UCase("RprHis"), "Y", "N") + "'") + _
          "Join Fgd on FdCoCd= FmdCoCd and FdRefYy= FmdFdYy and FdKey= FmdFdKey " + IIF(gs_Partition = ctCurrPrtn, " And FdPrtKey=FmdPrtKey ", "") + _
          wCndFg + IIF(wAddlCnd <> "", IIF(wCndFg <> "", " And ", " Where ") + wAddlCnd, IIF(gs_Partition = ctCurrPrtn, IIF(wCndFg <> "", " And ", " Where ") + " FmdPrtKey='" + ctCurrPrtn + "' ", ""))
  '****** Sachin 2.13.0 - 24/06/2006 ******
'Repair Jri,Jro change-JRI JRO records removed from detail section, wSel6, wLnk6,wSel7, wLnk7 are removed
'  'repair -field selection
'  wSel6 = " OmDt as TDt, OdTc as TdTc, OdYy as TdYy, OdChr as TdChr, OdNo as TdNo, OdSr as TdSr, 0 as TdSrNo, " + _
'          "OdYy as qTdYy,case when OdTc = 'JRI' then 0 else 999999 end as TdKey, '' as qFdYy, 0 as FdKey,'' as qFmdYy,0 as FmdKey, " + _
'          "OrdDsg.ModUsr as qUsrCd, OrdDsg.ModTime as qModTime, '' as TdFrBLoc, '' as TdWrk, '' as TdToBLoc, " + _
'          "'' as TdByWrk, IsNull(RmCtg, '') as RmCtg, IsNull(OrRmCd,'') as qRmCd,'' as qLotNo, " + _
'          "IsNull(OrLn1,0) as qRmSz, IsNull(OrLn2,0) as qRmSz2, IsNull(OrLn3,0) as qRmSz3, " + _
'          wRmSzDesc + " as qRmSzDesc,  0 as qStkRt, " + _
'          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
'          "IsNull(OrQty*OdOrdQty,0) as qRmQty, IsNull(OrWt*OdOrdQty,0) as qRmWt,case when OdTc='JRI' then 1 else 0 end as qIn," + _
'          "case when OdTc='JRO' then 1 else 0 end as qOut, " + _
'          "0 as qPcsIn, 0 as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, " + _
'          "0 as qPcsRej, 0 as qPcsPtRej, 0 as qIssQty, 0 as qIssPtQty, 0 as qIssWt, " + _
'          "'' as qLossRmCd, 0 as qLossRmWt, '' as qOthBYy, '' as qOthBChr, " + _
'          "0 as qOthBNo, '*' as qFlag, " + IIF(UCase(adc.MenuCd) <> UCase("RprHis"), "BFgSubLoc", "''") + " as BFgSubLoc "
'
'  'wLnk6 = " from Bag " + ws_BagJoin + _
'          "Join OrdDsg Od on OdDmStkYy =BYy and OdDmStkChr =BCHR and OdDmStkNo =BNo " + _
'          "Join OrdMst on OmCoCd= Od.OdCoCd and OmTc= od.OdTc and OmYy= Od.OdYy and OmChr= Od.OdChr and OmNo= Od.OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=Od.OdPrtKey ", "") + _
'          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= OmTc " + _
'          "where " + IIF(gs_Partition = ctCurrPrtn, " Od.OdPrtKey='" + ctCurrPrtn + "'", "") + IIF(wAddlCnd <> "", " and " + wAddlCnd, "") + wCnd
'  'repair - link
'  wLnk6 = " from Bag " + ws_BagJoin + _
'          ws_OdBagJoin + _
'          "Join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and " + _
'          "OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
'          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= OmTc " + _
'          "Left outer join OrdRm on OrCoCd = OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr " + _
'          "and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + _
'          "Left outer join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + _
'          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OrLn1 " + _
'          "where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "'", "") + _
'          IIF(wAddlCnd <> "", " and " + wAddlCnd, "") + wCnd
'          '"Join OrdDsg on OdDmStkYy =BYy and OdDmStkChr =BCHR and OdDmStkNo =BNo and OdCoCd= BCoCd" + _
'          IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=BPrtKey ", "")
'  wSel7 = " OmDt as TDt, OdTc as TdTc, OdYy as TdYy, OdChr as TdChr, OdNo as TdNo, OdSr as TdSr, 0 as TdSrNo, " + _
'          "OdYy as qTdYy,case when OdTc = 'JRI' then 0 else 999999 end as TdKey, '' as qFdYy, 0 as FdKey,'' as qFmdYy,0 as FmdKey, " + _
'          "OrdDsg.ModUsr as qUsrCd, OrdDsg.ModTime as qModTime, '' as TdFrBLoc, '' as TdWrk, '' as TdToBLoc, " + _
'          "'' as TdByWrk, IsNull(RmCtg, '') as RmCtg, IsNull(OrRmCd,'') as qRmCd,'' as qLotNo, " + _
'          "IsNull(OrLn1,0) as qRmSz, IsNull(OrLn2,0) as qRmSz2, IsNull(OrLn3,0) as qRmSz3, " + _
'          wRmSzDesc + " as qRmSzDesc,  0 as qStkRt, " + _
'          "IsNull((Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg), '') as qPtrYN, " + _
'          "IsNull(OrQty*OdOrdQty,0) as qRmQty, IsNull(OrWt*OdOrdQty,0) as qRmWt,case when OdTc='JRI' then 1 else 0 end as qIn," + _
'          "case when OdTc='JRO' then 1 else 0 end as qOut, " + _
'          "0 as qPcsIn, 0 as qPrtIn, 0 as qPcsOut, 0 as qPrtOut, " + _
'          "0 as qPcsRej, 0 as qPcsPtRej, 0 as qIssQty, 0 as qIssPtQty, 0 as qIssWt, " + _
'          "'' as qLossRmCd, 0 as qLossRmWt, '' as qOthBYy, '' as qOthBChr, " + _
'          "0 as qOthBNo, '*' as qFlag, " + IIF(UCase(adc.MenuCd) <> UCase("RprHis"), "BFgSubLoc", "''") + " as BFgSubLoc "
'  wLnk7 = " from Bag " + ws_BagJoin + _
'          "Join OrdDsg on OdDmStkYy =BYy and OdDmStkChr =BCHR and OdDmStkNo =BNo and OdCoCd= BCoCd And OdPrtKey=BPrtKey " + _
'          "Join OrdMst on OmTc ='JRO' and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and " + _
'          "OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
'          "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= OmTc " + _
'          "Left outer join OrdRm on OrCoCd = OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr " + _
'          "and OrNo= OdNo and OrSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " And OrPrtKey=OdPrtKey ", "") + _
'          "Left outer join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=OrPrtKey ", "") + _
'          "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OrLn1 " + _
'          "where " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey='" + ctCurrPrtn + "'", "") + _
'          IIF(wAddlCnd <> "", " and " + wAddlCnd, "") + wCnd + _
'          " and OMDT >= (select Om.OmDt from OrdMST Om where Om.OMTc ='JRI' AND Om.OMYy = BRepOdYy And Om.OMChr= BRepOdChr And Om.OMNo = BRepOdNo )"
              
  wSqlStrg = "Select " + wCommSel + wSel1 + wLnk1 + " Union All " + _
             "Select " + wCommSel + wSel2 + wLnk2 + " Union All " + _
             "Select " + wCommSel + wSel3 + wLnk3 + " Union All " + _
             "Select " + wCommSel + wSel4 + wLnk4
    
'Repair Jri,Jro change-JRI JRO records removed from detail section
'  'repair - query added
'  If UCase(adc.MenuCd) = UCase("RprHis") Then _
    wSqlStrg = wSqlStrg + " Union All " + _
                  "Select " + wCommSel + wSel6 + wLnk6 + " Union All " + _
                  "Select " + wCommSel + wSel7 + wLnk7
             

    If UCase(adc("UoPMCdTo")) = "3D" Then GRP_REP.SetFormula Rep, "wPicOpt", "'3D'"
    If UCase(adc("UoPMCdTo")) = "HSK" Then GRP_REP.SetFormula Rep, "wPicOpt", "'HSK'"
    If UCase(adc("UoPMCdTo")) = "" Then GRP_REP.SetFormula Rep, "wPicOpt", "''": Rep.Pic.Suppress = True

  'Debug.Print "Select " + wCommSel + wSel3 + wLnk3 + ""
  '****** Sachin 2.13.0 - 25/06/2006 - Optional viewing of FM transactions ******
  If adc("UoYn3") = "Y" Then _
    wSqlStrg = wSqlStrg + " Union All " + _
             "Select " + wCommSel + wSel5 + wLnk5
  '****** Sachin 2.13.0 - 25/06/2006 - Optional viewing of FM transactions ******
    
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 ctMaxDopOpt added
  
  Dim wSubRepSql As String, wSubRepStr1 As String, wSubRepGrp As String, wSubRepSrt As String
  'wSubRepGrp , wSubRepSrt
  
  Select Case UCase(adc("UoPMCdFr"))
  Case Is = "RMSCTG"
  'Repair Jri,Jro change-JRI JRO records removed from subreport group by
'    If UCase(adc.MenuCd) = UCase("RprHis") Then
'      wSubRepGrp = " group by Z.BYy, Z.BChr, Z.BNo, Z.RmCtg, Z.RmSCtg "
'      wSubRepSrt = " order by BYy, BChr, BNo, qSrt, RmSCtg "
'    Else
      wSubRepGrp = " group by BYy, BChr, BNo, RmCtg, RmSCtg "
      wSubRepSrt = " order by BYy, BChr, BNo, qSrt, RmSCtg "
      wSubRepStr1 = " '*' as TdRmCd "
'    End If
  Case Is = "RMCD"
  'Repair Jri,Jro change-JRI JRO records removed from subreport group by
'    If UCase(adc.MenuCd) = UCase("RprHis") Then
'      wSubRepGrp = " group by Z.BYy, Z.BChr, Z.BNo, Z.RmCtg, Z.RmSCtg, Z.TdRmCd "
'      wSubRepSrt = " order by BYy, BChr, BNo, qSrt, RmSCtg, TdRmCd "
'    Else
      wSubRepGrp = " group by BYy, BChr, BNo, RmCtg, RmSCtg, TdRmCd "
      wSubRepSrt = " order by BYy, BChr, BNo, qSrt, RmSCtg, TdRmCd "
      wSubRepStr1 = " TdRmCd "
'    End If
  End Select
       
  If UCase(adc("UoPMCdFr")) <> "N" Then
    ' **** Zubin 211 **** ' added ws_BagJoin
    ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' included
    If UCase(adc.MenuCd) <> UCase("RprHis") Then
    wSubRepSql = "Select BYy, BChr, BNo, RmCtg, RmSCtg, " + wSubRepStr1 + ", " + _
                 "sum(TdRmQty* (Case when LFr.LocTyp= 'W' and TdFrRmDc='D' then 1 " + _
                 "when LFr.LocTyp= 'W' and TdFrRmDc='C' then -1 " + _
                 "when LTo.LocTyp= 'W' and TdToRmDc='D' then 1 " + _
                 "when LTo.LocTyp= 'W' and TdToRmDc='C' then -1 Else 0 end)) as TdRmQty, " + _
                 "sum(TdRmWt* (Case when LFr.LocTyp= 'W' and TdFrRmDc='D' then 1 " + _
                 "when LFr.LocTyp= 'W' and TdFrRmDc='C' then -1 " + _
                 "when LTo.LocTyp= 'W' and TdToRmDc='D' then 1 " + _
                 "when LTo.LocTyp= 'W' and TdToRmDc='C' then -1 Else 0 end)) as TdRmWt, " + _
                 "(case RmCtg when 'D' then 1 when 'C' then 2 when 'G' then 3 " + _
                 "when 'P' then 4 When 'S' then 5 When 'L' then 9 when 'X' then 6 when 'A' then 7 " + _
                 "when 'M' then 8 else 0 end) as qSrt " + _
                 "from Bag " + ws_BagJoin + " join Txnd on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and " + _
                 "TdBNo= BNo and TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey=BPrtKey ", "") + _
                 "join RmMst on RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TdPrtKey ", "") + _
                 "join Loc LFr on LFr.LocCoCd= TdCoCd and LFr.LocCd= TdFrRmLoc " + _
                 "join Loc LTo on LTo.LocCoCd= TdCoCd and LTo.LocCd= TdToRmLoc " + _
                 "join OrdMst on OmCoCd= BCoCd and OmTc= BOdTc and OmYy= BOdYy and OmChr= BOdChr and " + _
                 "OmNo= BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey And BPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + wSubRepGrp + wSubRepSrt
    ' **** Zubin 211 **** ' added ws_BagJoin
    Else
    'Repair Jri,Jro change-JRI JRO records common selection removed from subreport
'    wSubRepSql = " SELECT Z.BYy AS BYy, Z.BChr AS BChr, Z.BNo AS BNo, Z.RmCtg AS RmCtg, Z.RmSCtg AS RmSCtg, (case when '" + adc("UoPMCdFr") + "'='RMCD' then Z.TdRmCd else '*' end) AS TdRmCd," + _
                 "SUM(Z.TdRmQty) AS TdRmQty,SUM(Z.TdRmWt) AS TdRmWt,max(Z.qSrt) as qSrt from " + _
                "(Select BYy, BChr, BNo, RmCtg, RmSCtg, (case when '" + adc("UoPMCdFr") + "'='RMCD' then TdRmCd else '*' end) AS TdRmCd, "
    wSubRepSql = "Select BYy, BChr, BNo, RmCtg, RmSCtg, " + wSubRepStr1 + ", " + _
                 "SUM(TdRmQty* (Case when LFr.LocTyp= 'XR' and TdFrRmDc='D' then 1 " + _
                 "when LFr.LocTyp= 'XR' and TdFrRmDc='C' then -1 " + _
                 "when LTo.LocTyp= 'XR' and TdToRmDc='D' then 1 " + _
                 "when LTo.LocTyp= 'XR' and TdToRmDc='C' then -1 Else 0 end)) as TdRmQty, " + _
                 "SUM(TdRmWt* (Case when LFr.LocTyp= 'XR' and TdFrRmDc='D' then 1 " + _
                 "when LFr.LocTyp= 'XR' and TdFrRmDc='C' then -1 " + _
                 "when LTo.LocTyp= 'XR' and TdToRmDc='D' then 1 " + _
                 "when LTo.LocTyp= 'XR' and TdToRmDc='C' then -1 Else 0 end)) as TdRmWt, " + _
                 "(case RmCtg when 'D' then 1 when 'C' then 2 when 'G' then 3 when 'P' then 4 When 'S' then 5 " + _
                 "When 'L' then 9 when 'X' then 6 when 'A' then 7 when 'M' then 8 else 0 end) as qSrt " + _
                 "from Bag " + ws_BagJoin + " join Txnd on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and " + _
                 "TdBNo= BNo and TdSrNo<> 0 " + IIF(gs_Partition = ctCurrPrtn, " And TdPrtKey=BPrtKey ", "") + _
                 "Join vParam ON vPTyp ='CHR' AND vPCoCd =TdCoCd AND vPMCd =TdTc AND vPSCd =TdChr AND vPValue3 ='Y' " + _
                 "join RmMst on RmCd= TdRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=TdPrtKey ", "") + _
                 "join Loc LFr on LFr.LocCoCd= TdCoCd and LFr.LocCd= TdFrRmLoc " + _
                 "join Loc LTo on LTo.LocCoCd= TdCoCd and LTo.LocCd= TdToRmLoc " + _
                 "join OrdMst on OmCoCd= BCoCd and OmTc= BRepOdTc and OmYy= BRepOdYy and OmChr= BRepOdChr and " + _
                 "OmNo= BRepOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey And BPrtKey='" + ctCurrPrtn + "' ", "") + _
                 "Join Txn on TDt >=OmDt  and TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo And TPrtKey=TdPrtKey " + wCnd + wSubRepGrp + wSubRepSrt
    
    'Repair Jri,Jro change-JRI JRO records removed from subreport
'    wSubRepSql = wSubRepSql + " Union All Select  BYy, BChr, BNo,IsNull(RmCtg, '') as RmCtg,IsNull(RmSCtg, '') as RmSCtg, (case when '" + adc("UoPMCdFr") + "'='RMCD' then IsNull(OrRmCd,'') else '*' end) as TdRmCd," + _
                "IsNull(OrQty*OdOrdQty,0) as TdRmQty, IsNull(OrWt*OdOrdQty,0) as TdRmWt,(case RmCtg when 'D' then 1 when 'C' then 2 when 'G' then 3 when 'P' then 4 When 'S' then 5 " + _
                "When 'L' then 9 when 'X' then 6 when 'A' then 7 when 'M' then 8 else 0 end) as qSrt from Bag " + ws_BagJoin + _
                " Join OrdDsg on OdCoCd= BCoCd and OdTc = BRepOdTc And OdYy = BRepOdYy And OdChr= BRepOdChr And OdNo = BRepOdNo and OdSr= BRepOdSr  And OdPrtKey=BPrtKey " + _
                "Join OrdMst on OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo  And OmPrtKey=OdPrtKey " + _
                "join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= OmTc " + _
                "Left outer join OrdRm on OrCoCd = OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr  And OrPrtKey=OdPrtKey " + _
                "Left outer join RmMst on RmCd= OrRmCd  And RmPrtKey=OrPrtKey Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OrLn1 " + _
                "where  OdPrtKey='C' " + wCnd + _
                " Union All Select  BYy, BChr, BNo,IsNull(RmCtg, '') as RmCtg,IsNull(RmSCtg, '') as RmSCtg, (case when '" + adc("UoPMCdFr") + "'='RMCD' then IsNull(OrRmCd,'') else '*' end) as TdRmCd," + _
                "(IsNull(OrQty*OdOrdQty,0)*-1) as TdRmQty,(IsNull(OrWt*OdOrdQty,0)*-1) as TdRmWt,(case RmCtg when 'D' then 1 when 'C' then 2 when 'G' then 3 when 'P' then 4 When 'S' then 5 " + _
                "When 'L' then 9 when 'X' then 6 when 'A' then 7 when 'M' then 8 else 0 end) as qSrt from Bag " + ws_BagJoin + _
                " Join OrdDsg on OdDmStkYy =BYy and OdDmStkChr =BCHR and OdDmStkNo =BNo and OdCoCd= BCoCd And OdPrtKey=BPrtKey " + _
                "Join OrdMst on OmTc ='JRO' and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo  And OmPrtKey=OdPrtKey " + _
                "Join Param Prm on Prm.PTyp= 'TC' and Prm.PMCd= OmTc Left outer join OrdRm on OrCoCd = OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr  And OrPrtKey=OdPrtKey " + _
                "Left outer join RmMst on RmCd= OrRmCd  And RmPrtKey=OrPrtKey Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OrLn1 " + _
                "where  OdPrtKey='C' " + wCnd + " and OMDT >= (select Om.OmDt from OrdMST Om where Om.OMTc ='JRI' AND Om.OMYy = BRepOdYy And Om.OMChr= BRepOdChr And Om.OMNo = BRepOdNo )" + _
                " ) AS Z " + wSubRepGrp + wSubRepSrt
    End If
    Rep.Subreport1.Suppress = False
  ElseIf UCase(adc("UoPMCdFr")) = "N" Then
    wSubRepSql = "Select ' ' as BYy, ' ' as BChr, 0 as BNo, ' ' as RmCtg, ' ' as RmSCtg, ' ' as TdRmCd, 0 as TdRmQty, 0 as TdRmWt, 0 as qSrt "
    Rep.Subreport1.Suppress = True
  End If
  
  Call moCn.RepRes(Rep.Subreport1.OpenSubreport, wSubRepSql + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added

  '3.11.0
  If adc("UoYN") = "Y" Then
    Rep.TxtSzStkRt.SetText "Stk Rt"
    Rep.FldRmStkRt.Suppress = False
    'Rep.FldRmSz.Suppress = True
  Else
    Rep.TxtSzStkRt.SetText ""
    Rep.FldRmStkRt.Suppress = True
    'Rep.FldRmSz.Suppress = False
  End If
  'Repair Jri,Jro change-Repair GrsWt label & textbox will be suppress for non repair bag
  If UCase(adc.MenuCd) = UCase("RprHis") Then
    Rep.hGrsWtIn.Suppress = False
    Rep.FldGrsWtIn.Suppress = False
  Else
    Rep.hGrsWtIn.Suppress = True
    Rep.FldGrsWtIn.Suppress = True
  End If
 
 
  '*** Jay 2.13(CT) ***
  If adc("UoBagTyp") = "F" Then
    Rep.TxtBagTypAndChr.SetText "(Flute Bags)"
  ElseIf adc("UoBagTyp") = "P" Then
    Rep.TxtBagTypAndChr.SetText "(Primary Bags)"
  ElseIf adc("UoBagTyp") = "C" Then
    Rep.TxtBagTypAndChr.SetText "(Component Bags)"
  End If
  ''***(Jen 2.11)
  'If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & ")"
  ''***(Jen 2.11)
  '*** Jay 2.13(CT) ***
  
  ' ****** Manali 3.03 - 08/07/08 - RmSzDesc
'  GRP_REP.SetFormula Rep, "wRmSzPtr", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} " + _
                                      "Else (If {rdo.qPtrYn} = 'Y' Then ToText({rdo.qRmSz}, 4) Else ToText({rdo.qRmSz}, 3))"
  GRP_REP.SetFormula Rep, "qRmSz", "If {rdo.qRmSzDesc} <> '' Then {rdo.qRmSzDesc} Else ''"
                                       
  ' ****** Manali 3.03 - 08/07/08 - RmSzDesc
    
  'Sachin - 4.1.0.0
  If adc("UOYN1") = "N" Then
    Rep.TxtOthDetHd.SetText "Other Bag"
    GRP_REP.SetFormula Rep, "wOthBag", "If {rdo.qOthBNo}> 0 Then Trim ({rdo.qOthBYy})+' / '+Trim ({rdo.qOthBChr})+' / '+Trim (ToText ({rdo.qOthBNo}, '######0'))"
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.qRmCd}", 5.5)
  SetSubFormula Rep.Subreport1, "wRmCdFontSzSumm", GetRmcdFontSz("{ado.TdRmCd}", 7)
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub

' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 3
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
    
  '3.11.0
  Case Is = 0
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
    
  Case Is = 1   'Sachin - 4.1.0.0
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2   '6.1b-41
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added

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

  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = True
  'Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoOmTcFr"))    '****** Before 3.02 ******
  adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
    
  Set adc.FirNKeyCtl = adc("UoOmTcFr")   ' **** Manali 3.5.0 - 17/11/08
  Call InitProp(Me)
    
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gs_LocTyp = "'P'"
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"
  gs_CmCtg = "C"
  gs_PTyp = "BHISTOPT"
  gs_Tbl = "Param"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  '4.1.0.0
  ms_BHDets = moCn.GetFldVal("Select hShowBgHistDetsYN From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + gs_CoCd + "'")
  If Trim(ms_BHDets) = "Y" Then
    ChkBoxArr(1).Visible = True
  Else
    ChkBoxArr(1).Visible = False
  End If
  If UCase(adc.MenuCd) = UCase("RprHis") Then ChkBoxArr(2).Visible = True    '6.1b-41 complete history option is visible only for repair bag history
  
  Call SetGroupSort
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  gb_CoCdFor = True       '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
    gs_Tbl = "Param"
    gs_PTyp = "BHISTOPT"
    
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Case Is = UCase("UoCoCdFr")
  If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
     Me.Caption = GetFrmCaption(pv_NewValue, "Bag History")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  
  Case Is = UCase("UoPmCdFr")
    '*** Jay 2.14 *** (SEO Change)
    If pv_NewValue <> "Y" Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("Select IsNull ((Select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            " Where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            " UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            " Isnull ((Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            Cancel = True: ErrMsg = "Enter SEO Password to View Stock Rate ": Exit Sub
    End If
    '**Bhavna added pic opt
   Case Is = UCase("UoPmCdTo")
    
    If pv_NewValue <> "" And pv_NewValue <> "HSK" And pv_NewValue <> "3D" Then
       Cancel = True: ErrMsg = "Enter 3D/HSK/Blank View Picture ": Exit Sub
    End If
 
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)

  gb_CoCdFor = True       '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  'indigo.164 sketch design tctype and order tc added
  gs_DmTcTyp = "DM','SM"
  gs_OmTcTyp = "SO','OS"

  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
    gs_Tbl = "Param"
    gs_PTyp = "BHISTOPT"
  
  Case Is = UCase("UOOMTCFR")
          If UCase(adc.MenuCd) = UCase("RprHis") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
  Case Is = UCase("UOOMTCTO")
          If UCase(adc.MenuCd) = UCase("RprHis") Then HlpList.PMCd "TC", "'JT'":  Exit Sub
          
  Case Is = UCase("UOPMCDTO")
    gs_Tbl = "Param"
    gs_PTyp = "PCT"
    
  End Select
  
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
 '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + ctSelfCoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    adc("UoBagTyp") = "P"
  End If
  
  '*** Jay 2.13(CT) ***
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), "Bag History")
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******

'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblPrimFluteBags.Visible = False
'    adc("UoBagTyp").Visible = False
'    adc("UoBagTyp") = "N"
'  End If
'  adc("UOYN3") = "N"  '****** Sachin 2.13.0 - 25/06/2006 ******
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***


' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added
If adc("UoYN3") = "" Then adc("UoYN3") = "N"
If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
' ***** Manali 3.5.0 - 24/11/08 - Check Boxes Added

'Sachin - 4.1.0.0
adc("UOYN1") = IIF(Trim(ms_BHDets) = "N", "N", "Y")
If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked

adc("UoYN") = "N"     '3.11.0
adc("UoYN2") = "N"     '6.1b-41 default value for complete bag history is false
ChkBoxArr(2).Value = Unchecked  '6.1b-41 default value for complete bag history

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                            "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName

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

Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(0).Value = Unchecked
    End If
End Sub
