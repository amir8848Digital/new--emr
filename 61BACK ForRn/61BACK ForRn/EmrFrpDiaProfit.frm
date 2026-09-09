VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpDiaProfit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diamond Profitability Report"
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   52
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
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11070
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5055
      TabIndex        =   60
      Top             =   9615
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   61
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
         TabIndex        =   62
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
      Left            =   6270
      TabIndex        =   59
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   56
         Top             =   360
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
      Height          =   9720
      Left            =   0
      TabIndex        =   55
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   54
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpDiaProfit.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpDiaProfit.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   58
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Width           =   8865
               _ExtentX        =   15637
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   57
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Stock Rate "
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
               Index           =   0
               Left            =   4200
               TabIndex        =   49
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   5670
               Width           =   2220
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   48
               Top             =   5670
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4260
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2745
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "IrRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6960
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2745
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "IrRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4260
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2460
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
               Index           =   2
               Left            =   6960
               TabIndex        =   19
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2460
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
               Index           =   3
               Left            =   1560
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   1890
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1560
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2460
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "RmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1560
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2745
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "IrRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4260
               TabIndex        =   32
               ToolTipText     =   "Enter To Design Code"
               Top             =   3930
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "IdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6960
               TabIndex        =   33
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   3930
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "IdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1560
               TabIndex        =   31
               ToolTipText     =   "Enter From Design Code"
               Top             =   3930
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "IdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   5055
               TabIndex        =   38
               ToolTipText     =   "Enter To Order Year"
               Top             =   4215
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "IdExpOdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2355
               TabIndex        =   35
               ToolTipText     =   "Enter From Order Year"
               Top             =   4215
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "IdExpOdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1560
               TabIndex        =   34
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "IdExpOdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2805
               TabIndex        =   36
               ToolTipText     =   "Enter From Order Character"
               Top             =   4215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "IdExpOdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4260
               TabIndex        =   37
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "IdExpOdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5505
               TabIndex        =   39
               ToolTipText     =   "Enter To Order Character"
               Top             =   4215
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "IdExpOdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5175
               TabIndex        =   43
               ToolTipText     =   "Enter To Order Serial"
               Top             =   4500
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "IdExpOdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2475
               TabIndex        =   41
               ToolTipText     =   "Enter From Order Serial"
               Top             =   4500
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "IdExpOdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1560
               TabIndex        =   40
               ToolTipText     =   "Enter From Order Number"
               Top             =   4500
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "IdExpOdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4260
               TabIndex        =   42
               ToolTipText     =   "Enter To Order Number"
               Top             =   4500
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "IdExpOdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   5040
               TabIndex        =   4
               ToolTipText     =   "Enter To Invoice Year"
               Top             =   750
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "IdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   2340
               TabIndex        =   1
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   750
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1560
               TabIndex        =   0
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   750
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2760
               TabIndex        =   2
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   750
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4260
               TabIndex        =   3
               ToolTipText     =   "Enter To Invoice Tc"
               Top             =   750
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "IdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5460
               TabIndex        =   5
               ToolTipText     =   "Enter To Invoice Character"
               Top             =   750
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "IdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1560
               TabIndex        =   6
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   1035
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   4260
               TabIndex        =   7
               ToolTipText     =   "Enter To Invoice Number"
               Top             =   1035
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "IdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4260
               TabIndex        =   11
               ToolTipText     =   "Enter To Customer"
               Top             =   1605
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "InCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6960
               TabIndex        =   12
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1605
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "InCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   1560
               TabIndex        =   44
               ToolTipText     =   "Specify Whether Base Value Is Actual, Invoice Or Customs "
               Top             =   4785
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   1560
               TabIndex        =   46
               ToolTipText     =   "Specify Whether Other Value Is Actual, Invoice Or Customs "
               Top             =   5070
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL1"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4260
               TabIndex        =   9
               ToolTipText     =   "Enter To Invoice Date"
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "InDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1560
               TabIndex        =   10
               ToolTipText     =   "Enter From Customer"
               Top             =   1605
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "InCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   1560
               TabIndex        =   8
               ToolTipText     =   "Enter From Invoice Date"
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "InDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1560
               TabIndex        =   88
               ToolTipText     =   "Enter Company Code"
               Top             =   465
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "InCoCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1560
               TabIndex        =   47
               ToolTipText     =   "Enter Currency Code Whose Invoices You Wish To View Or Leave It Blank If You Wish To View All Invoices In Base Currency"
               Top             =   5355
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL,UOINVAL1"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1560
               TabIndex        =   28
               ToolTipText     =   "Enter From Lot No."
               Top             =   3630
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
               CmpStr          =   "IrLotNo>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4260
               TabIndex        =   29
               ToolTipText     =   "Enter To Lot No."
               Top             =   3630
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
               CmpStr          =   "IrLotNo<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6960
               TabIndex        =   30
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3630
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
               CmpStr          =   "IrLotNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   4260
               TabIndex        =   45
               ToolTipText     =   "Specify Whether Base Value Is Actual, Invoice Or Customs "
               Top             =   4800
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL2"
               ReCalcParent    =   "UOCURCDFR"
               IdName          =   "UOINVAL2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   1560
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Size From Range"
               Top             =   3030
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               Mask            =   "##0.0000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "UORMWTFR"
               IdName          =   "UORMWTFR"
               CmpStr          =   "IrRmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   4260
               TabIndex        =   24
               ToolTipText     =   "Enter Raw Material Size To Range"
               Top             =   3030
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               Mask            =   "##0.0000"
               MaxLength       =   8
               DataType        =   2
               DataField       =   "UORMWTTO"
               IdName          =   "UORMWTTO"
               CmpStr          =   "IrRmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   4260
               TabIndex        =   26
               ToolTipText     =   "Enter Raw Material Parent Code To Range"
               Top             =   3315
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
               CmpStr          =   "RmParentCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   1560
               TabIndex        =   25
               ToolTipText     =   "Enter Raw Material Parent Code From Range"
               Top             =   3315
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
               CmpStr          =   "RmParentCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   6960
               TabIndex        =   27
               ToolTipText     =   "Enter Raw Material Parent Code Selection"
               Top             =   3315
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   71
               Left            =   6600
               TabIndex        =   95
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   5640
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   4260
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2175
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "ISNULL(SC.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1560
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2175
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "ISNULL(SC.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6960
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2175
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "ISNULL(SC.PSCd,'') In"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stone Class"
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
               Left            =   30
               TabIndex        =   97
               Top             =   2175
               Width           =   1305
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
               Index           =   39
               Left            =   30
               TabIndex        =   96
               Top             =   5670
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Code"
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
               Left            =   30
               TabIndex        =   94
               Top             =   3315
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Size"
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
               Left            =   30
               TabIndex        =   93
               Top             =   3030
               Width           =   1425
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Oth Value As"
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
               Left            =   2760
               TabIndex        =   92
               Top             =   4800
               Width           =   1575
            End
            Begin VB.Label LblLotNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot No."
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
               Left            =   30
               TabIndex        =   91
               Top             =   3630
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
               Left            =   30
               TabIndex        =   90
               Top             =   5355
               Width           =   1005
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
               Index           =   29
               Left            =   30
               TabIndex        =   89
               Top             =   465
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Dt"
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
               Left            =   30
               TabIndex        =   87
               Top             =   1320
               Width           =   1425
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Value As"
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
               Left            =   30
               TabIndex        =   86
               Top             =   4785
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Other Value As"
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
               Left            =   30
               TabIndex        =   85
               Top             =   5070
               Width           =   1575
            End
            Begin VB.Label ALBL 
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
               Index           =   23
               Left            =   30
               TabIndex        =   84
               Top             =   750
               Width           =   1725
            End
            Begin VB.Label ALBL 
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
               Index           =   20
               Left            =   30
               TabIndex        =   83
               Top             =   1035
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
               Left            =   5400
               TabIndex        =   82
               Top             =   750
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
               Index           =   12
               Left            =   2700
               TabIndex        =   81
               Top             =   750
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
               Index           =   11
               Left            =   4980
               TabIndex        =   80
               Top             =   750
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
               Index           =   10
               Left            =   2280
               TabIndex        =   79
               Top             =   750
               Width           =   105
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
               Index           =   7
               Left            =   30
               TabIndex        =   78
               Top             =   1605
               Width           =   1425
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
               Left            =   30
               TabIndex        =   77
               Top             =   4215
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Ord No/Sr"
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
               Left            =   30
               TabIndex        =   76
               Top             =   4500
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
               Index           =   24
               Left            =   5415
               TabIndex        =   75
               Top             =   4215
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
               Left            =   2715
               TabIndex        =   74
               Top             =   4215
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
               Left            =   4965
               TabIndex        =   73
               Top             =   4215
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
               Left            =   2265
               TabIndex        =   72
               Top             =   4215
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
               Left            =   5085
               TabIndex        =   71
               Top             =   4500
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
               Index           =   31
               Left            =   2385
               TabIndex        =   70
               Top             =   4500
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
               Left            =   30
               TabIndex        =   69
               Top             =   3930
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   4
               Left            =   30
               TabIndex        =   68
               ToolTipText     =   "Location"
               Top             =   1890
               Width           =   1125
            End
            Begin VB.Label ALBL 
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
               Index           =   0
               Left            =   30
               TabIndex        =   67
               Top             =   2460
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   8
               Left            =   30
               TabIndex        =   66
               Top             =   2745
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
               Left            =   6960
               TabIndex        =   65
               Top             =   60
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
               Left            =   4260
               TabIndex        =   64
               Top             =   60
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
               Left            =   1560
               TabIndex        =   63
               Top             =   60
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpDiaProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepDiaProfit
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Rm Code", "RmSCtg+IrRmCd", "'('+ RmSCtg+ ') '+ IrRmCd", "RmSCtg,IrRmCd", "hRmSCtg,hIrRmCd", "Rm SubCtg", "", "", "RmDesc "
  GRP_REP.Add "Rm SubCtg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
  '*********Bhavna
'  GRP_REP.Add "Rm Size", "convert(varchar(8),IrRmSz)", "", "IrRmSz", "hRmSz", "", "", "", "convert(varchar(8),IrRmSz)"
  GRP_REP.Add "Rm Size", GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3"), "", "IrRmSz", "hRmSz", "", "", "", GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3")
  GRP_REP.Add "Rm Parent Cd", "RmParentCd", "", "RmParentCd", "hRmParentCd", "Rm Parent Cd", "", "", "(Select PDesc From Param where " + _
                                    "PTyp= 'PARENTCD' and PmCd = RmParentCd and PValue= RmCtg) "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Rm Customer", "RmCmCd", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else RmCmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  GRP_REP.Add "Inv Vch No", "IdTc+IdYy+IdChr+str(IdNo)", "IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))", "", "", "Inv Customer", "", "", "InExpNo "
  '6.1
  GRP_REP.Add "Inv Date", "convert(varchar(10), InDt, 112)", "convert(varchar(8), InDt, " + CStr(DtFmtSlahVal()) + ")", "InDt", "hInDt", "", "", "", ""
  GRP_REP.Add "Inv Customer", "InCmCd", "", "InCmCd", "hInCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= InCmCd) "
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Design Code", "IdDmCd", "", "IdDmCd", "hIdDmCd", "Design Ctg", "", "", "DmDesc"
  '6.1
  GRP_REP.Add "OrderNo", "IdExpOdTc+IdExpOdYy+IdExpOdChr+str(IdExpOdNo)", "IdExpOdTc+'/'+IdExpOdYy+'/'+IdExpOdChr+'/'+LTrim(Str(IdExpOdNo))", "", "", "", "", "", "'('+convert(varchar(8), OmDt, " + CStr(DtFmtDashVal()) + ") +'/'+ IdExpCmCd+ '/'+ OmPONo+ ')' "
  'sv.52 Stone Class group sort added
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(SC.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(SC.PSCd, '')", "", "", "", "", "", "SC.PDesc"
    
  If moCn.RecSeek("Select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + gs_CoCd + "' and HLotNoYN='Y'") Then _
    GRP_REP.Add "Rm Lot No", "IrLotNo", "", "", "", "", "", "", "IrLotNo"
  
'  GRP_REP.Add "Rm Code", "RmSCtg+IrRmCd", "'('+ RmSCtg+ ') '+ IrRmCd", "RmSCtg,IrRmCd", "hRmSCtg,hIrRmCd", "Rm SubCtg", "", "", "RmDesc "
'  GRP_REP.Add "Rm SubCtg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
'  GRP_REP.Add "Rm Customer", "RmCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
'  GRP_REP.Add "Stk Rate", "str(IrRmStkRt,9,2)", "", "IrRmStkRt", "hIrRmStkRt", "(RmCd) StkRt", "", "", ""
'  GRP_REP.Add "(RmCd) StkRt", "IrRmCd+str(IrRmStkRt,9,2)", "'('+IrRmCd+') '+LTrim(str(IrRmStkRt,9,2))", "RmSCtg,IrRmCd,IrRmStkRt", "hRmSCtg,hIrRmCd,hIrRmStkRt", "Rm Code,Rm SubCtg,Stk Rate", "", "", " '('+ RmSCtg+ ')'"
'  GRP_REP.Add "Inv Vch No", "IdTc+IdYy+IdChr+str(IdNo)", "IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))", "", "", "Inv Customer,Inv Exp No", "", "", "InExpNo "
'  GRP_REP.Add "Inv Exp No", "InExpNo", "", "", "", "Inv Customer", "", "", "IdTc+'/'+IdYy+'/'+IdChr+'/'+LTrim(Str(IdNo))"
'  GRP_REP.Add "Inv Date", "convert(varchar(10), InDt, 112)", "convert(varchar(8), InDt, 3)", "InDt", "hInDt", "Quarter,Month,Week", "", "", ""
'  GRP_REP.Add "Inv Customer", "InCmCd", "", "InCmCd", "hInCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= InCmCd) "
'  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,InDt)))+' - '+DateName(WW,InDt)", "", "", "", "Quarter,Month", "", "", ""
'  GRP_REP.Add "Month", "Str(DatePart(yyyy,InDt))+Str(DatePart(mm,InDt))", "LTrim(Str(DatePart(yyyy,InDt)))+' '+DateName(MM,InDt)", "", "", "Quarter", "", "", ""
'  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,InDt)))+' - '+DateName(QQ,InDt)", "", "", "", "", "", "", ""
'  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "Sales Ctg", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
'  GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd= DmSalCtg) "
'  GRP_REP.Add "Design Code", "IdDmCd", "", "IdDmCd", "hIdDmCd", "Design Ctg,Sales Ctg,(DsgCd)Colour", "", "", "DmDesc"
'  GRP_REP.Add "(DsgCd)Colour", "IdDmCd+OdDmCol", "'('+IdDmCd+') '+OdDmCol", "IdDmCd", "hIdDmCd", "Design Ctg,Design Code,Design Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
'  GRP_REP.Add "Design Colour", "OdDmCol", "", "", "", "(DsgCd)Colour", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
'  GRP_REP.Add "Design Size", "IdDmSz", "", "IdDmSz", "hIdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= IdDmSz)"
'  GRP_REP.Add "Suffix", "IdSfx", "", "IdSfx", "hIdSfx", "", "", "", ""
'  GRP_REP.Add "Karat", "IdKt", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = IdKt) "
'  GRP_REP.Add "VaCtg", "IdVaCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = IdVaCtg) "
'  GRP_REP.Add "OrderNo", "IdExpOdTc+IdExpOdYy+IdExpOdChr+str(IdExpOdNo)", "IdExpOdTc+'/'+IdExpOdYy+'/'+IdExpOdChr+'/'+LTrim(Str(IdExpOdNo))", "", "", "Ord Customer,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ IdExpCmCd+ '/'+ OmPONo+ ')' "
'  GRP_REP.Add "Ord Customer", "IdExpCmCd", "", "IdExpCmCd", "hIdExpCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= IdExpCmCd) "
'  GRP_REP.Add "Customer PO", "IdExpCmCd+OmPONo", "'('+ IdExpCmCd+ ') '+ LTrim(OmPONo)", "", "", "Ord Customer,OrderNo", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ IdExpCmCd+ ')'+ IdExpOdTc+'/'+IdExpOdYy+'/'+IdExpOdChr+'/'+LTrim(Str(IdExpOdNo))"

End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  '*** Report Sql ***
  '    Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '    space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '    space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '    InExpNo, InDt, InCmCd, IdTc, IdYy, IdChr, IdNo, IdSr, IrSrNo, IdDmCd,
  '    IdSfx, IdDmSz, IdQty, IdExpCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr,
  '    IdExpOdNo, IdExpOdSr, IdKt, IdVaCtg, DmCtg, DmSalCtg, IrRmCd,
  '    RmSCtg, IrRmAVal as qVal1, IrRmAVal as qVal2,
  '    IrRmAWt As qWt1, IrRmAWt As qWt2
  '    From InvRm, InvDsg, DsgMst, InvHd, RmMst
  '    Where 1 = 2
  '*** Report Sql ***
  
  Dim grpflds As String, wSqlStrg As String
  Dim wDetPos As Integer, wNonePos As Integer, wrepcnd As String, wCnd As String
  Dim wBaseIrRmWt As String, wBaseIrSetVal As String, wBaseIrRmVal As String
  Dim wOthIrRmWt As String, wOthIrSetVal As String, wOthIrRmVal As String
  Dim wBase As String, wOth As String
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
  
  '*** To check if Detail is selected ***
'  Dim wGrp(6) As String, i As Integer
'  wGrp(1) = GRP_REP.Value(gltGroup1)
'  wGrp(2) = GRP_REP.Value(gltGroup2)
'  wGrp(3) = GRP_REP.Value(gltGroup3)
'  wDetPos = 0: wNonePos = 0
'  For i = 1 To 3
'    If wGrp(i) = "(Detail)" Then wDetPos = i
'    If wGrp(i) = "(None)" Then wNonePos = i: Exit For
'  Next i
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  Dim ws_CurJoin As String, ws_BaseCurCd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If adc("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_InBaseCur = False
    ws_CurJoin = " Join CustMst On CmCtg= 'C' and CmCd= InCmCd "
    wCnd = wCnd + " and CmCurCd= '" + adc("UoCurCdFr") + "' "
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If
  ' Zubin 213

  If UCase(adc("UoInVal")) = "A" Then
    wBase = "Actual"
    wBaseIrRmWt = "IrRmAWt"
    ' Zubin 213 (Bef 213)
    'wBaseIrSetVal = "IrSetAVal": wBaseIrRmVal = "IrRmAVal"
    ' Zubin 213
    wBaseIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
    Rep.TxtBase.SetText "Base = Actual"
    'Rep.TxtActualDefn.SetText "Actual Rm Val = Actual Wt * Stk Rt"
  ElseIf UCase(adc("UoInVal")) = "I" Then
    wBase = "Invoice"
    wBaseIrRmWt = "IrRmIWt"
    ' Zubin 213 (Bef 213)
    'wBaseIrSetVal = "IrSetIVal": wBaseIrRmVal = "IrRmIVal"
    ' Zubin 213
    wBaseIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
    Rep.TxtBase.SetText "Base = Invoice"
  ElseIf UCase(adc("UoInVal")) = "Z" Then
    wBase = "Custom"
    wBaseIrRmWt = "IrRmZWt"
    ' Zubin 213 (Bef 213)
    'wBaseIrSetVal = "IrSetZVal": wBaseIrRmVal = "IrRmZVal"
    ' Zubin 213
    wBaseIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
    wBaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
    Rep.TxtBase.SetText "Base = Customs"
  
  '******* urmi Invoice RmStkRt
  ElseIf UCase(adc("UoInVal")) = "S" Then
    wBase = "Stock"
    wBaseIrRmWt = "IrRmAWt"
    wBaseIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wBaseIrRmVal = "IrStkRt*IrRmAWt"
    Rep.TxtBase.SetText "Base = Stock"
    
    '4.1.4
    If Trim(adc("UoInVal2")) <> "" Then
        Dim wABaseIrRmVal As String, wIBaseIrRmVal As String, wZBaseIrRmVal As String
        wABaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
        wIBaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
        wZBaseIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
        wBaseIrRmVal = " (Case When RSCtg.PValue1='N' Then " + IIF(UCase(adc("UoInVal2")) = "A", wABaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "I", wIBaseIrRmVal, _
                                                               IIF(UCase(adc("UoInVal2")) = "Z", wZBaseIrRmVal, wBaseIrRmVal))) + " Else " + wBaseIrRmVal + " End)"
                                                               
        Rep.TxtBase.SetText IIF(adc("UoInVal2") = "I", "Base = Stk-Inv", IIF(adc("UoInVal2") = "A", "Base = Stk-Act", "Base = Stk-Cstm"))
    End If
  End If
  
  If UCase(adc("UoInVal1")) = "A" Then
    wOth = "Actual"
    wOthIrRmWt = "IrRmAWt"
    ' Zubin 213 (Bef 213)
    'wOthIrSetVal = "IrSetAVal": wOthIrRmVal = "IrRmAVal"
    ' Zubin 213
    wOthIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
    Rep.TxtOth.SetText "Other = Actual"
    Rep.TxtActualDefn.SetText "Actual"
  ElseIf UCase(adc("UoInVal1")) = "I" Then
    wOth = "Invoice"
    wOthIrRmWt = "IrRmIWt"
    ' Zubin 213 (Bef 213)
    'wOthIrSetVal = "IrSetIVal": wOthIrRmVal = "IrRmIVal"
    ' Zubin 213
    wOthIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end)", "IrRmIVal")
    Rep.TxtOth.SetText "Other = Invoice"
  ElseIf UCase(adc("UoInVal1")) = "Z" Then
    wOth = "Custom"
    wOthIrRmWt = "IrRmZWt"
    ' Zubin 213 (Bef 213)
    'wOthIrSetVal = "IrSetZVal": wOthIrRmVal = "IrRmZVal"
    ' Zubin 213
    wOthIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
    wOthIrRmVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
    Rep.TxtOth.SetText "Other = Customs"
  '******* urmi Invoice RmStkRt
  ElseIf UCase(adc("UoInVal1")) = "S" Then
    wOth = "Stock"
    wOthIrRmWt = "IrRmAWt"
    ' Zubin 213 (Bef 213)
    'wOthIrSetVal = "IrSetAVal"
    ' Zubin 213
    wOthIrSetVal = IIF(wb_InBaseCur, " (case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
    wOthIrRmVal = "IrStkRt*IrRmAWt"
    ' Zubin 213 (Err rectified)
    'Rep.TxtBase.SetText "Base = Stock"
    Rep.TxtOth.SetText "Other = Stock"
  End If
 
   '****** Sachin 3.02.0 - Additional Fields in Where Clause
    'wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " IrRmCtg >= '" + Trim(adc("UORMCTGFR")) + "'", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  '****** Sachin 3.02.0 - Additional Fields in Where Clause
 
  '**********bhavna added IrRmSz and RmParentCd
  ' Zubin 213 (ws_CurJoin added)
  '****** Sachin 3.02 - Additional Fields in Where Clause
  'sv.52 left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls added
  wSqlStrg = "Select " + grpflds + ", " + _
             "InExpNo, InDt, InCmCd, IdTc, IdYy, IdChr, IdNo, IdSr, IrSrNo, IrRmSz, IrRmSz2, IrRmSz3, IdDmCd, IdSfx, " + _
             "IdDmSz, IdQty, IdExpCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, " + _
             "IdExpOdSr, IdKt, IdVaCtg, DmCtg, DmSalCtg, IrRmCd, RmSCtg, RmParentCd, " + _
             wBaseIrRmVal + " as qVal1, " + wOthIrRmVal + " as qVal2, " + _
             wBaseIrRmWt + " as qWt1, " + wOthIrRmWt + " as qWt2 " + _
             "From InvRm join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy " + _
             "and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=IrPrtKey ", "") + _
             "join DsgMst on DmIdNo=IdDmIdno and DmTcTyp= 'DM' and DmCd= IdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
             "join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " And InPrtKey=IdPrtKey ", "") + _
             "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
             "and OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " And OdPrtKey=IdPrtKey ", "") + _
             "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=OdPrtKey ", "") + _
             ws_CurJoin + " join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " And RmPrtKey=IrPrtKey ", "") + _
             " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=RmCtg And RSCtg.PSCd=RmSCtg " + _
             " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
             " where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' And ", "") + " RmCtg in ('D', 'C') " + _
             wCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", "") + wAddlCnd, "")
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt   '****** Sachin 3.01 ctMaxDopOpt added
  
  If adc("UoRmCtgFr") <> "" Then
    Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
  Else
    Rep.TxtRmCtg.Suppress = True
    Rep.hOrRmCtg.Suppress = True
  End If
  Rep.h1Val1.SetText wBase + Chr(13) + "Rm Val": Rep.h2Val1.SetText wBase + Chr(13) + "Rm Val"
  Rep.h1Val2.SetText wOth + Chr(13) + "Rm Val": Rep.h2Val2.SetText wOth + Chr(13) + "Rm Val"
  Rep.h1Wt1.SetText wBase + Chr(13) + "Rm Wt": Rep.h2Wt1.SetText wBase + Chr(13) + "Rm Wt"
  Rep.h1Wt2.SetText wOth + Chr(13) + "Rm Wt": Rep.h2Wt2.SetText wOth + Chr(13) + "Rm Wt"
    
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If UCase(adc("UOINVAL")) = "S" Then
  If adc("UoYN") = "Y" Then
    Rep.FldqVal1.Suppress = False
    Rep.FldG1qVal1.Suppress = False
    Rep.FldG2qVal1.Suppress = False
    Rep.FldG3qVal1.Suppress = False
    Rep.FldGqVal1.Suppress = False
  Else
    Rep.h1Val1.SetText ""
    Rep.h2Val1.SetText ""
    Rep.FldqVal1.Suppress = True
    Rep.FldG1qVal1.Suppress = True
    Rep.FldG2qVal1.Suppress = True
    Rep.FldG3qVal1.Suppress = True
    Rep.FldGqVal1.Suppress = True
  End If
  End If
  If UCase(adc("UoInVal1")) = "S" Then
  If adc("UoYN") = "Y" Then
    Rep.FldqVal2.Suppress = False
    Rep.FldG1qVal2.Suppress = False
    Rep.FldG2qVal2.Suppress = False
    Rep.FldG3qVal2.Suppress = False
    Rep.FldGqVal2.Suppress = False
  Else
    Rep.h1Val2.SetText ""
    Rep.h2Val2.SetText ""
    Rep.FldqVal2.Suppress = True
    Rep.FldG1qVal2.Suppress = True
    Rep.FldG2qVal2.Suppress = True
    Rep.FldG3qVal2.Suppress = True
    Rep.FldGqVal2.Suppress = True
  End If
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.IrRmCd}", 8)

  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoInTcFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = True
  gb_RmZ = False
  gs_DmTcTyp = "DM"
  gs_OmTcTyp = "SO"
  gs_PTyp = "VACTG"
  gs_CmCtg = "C"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' **** Manali 350Nxt
  If Not moCn.RecSeek("Select 'x' from Head where HCd='" + ctSelfCmCd + "' and HCoCd='" + gs_CoCd + "' and HLotNoYN='Y'") Then
    LblLotNo.Visible = False
    adc("UoLotNoFr").Enabled = False: adc("UoLotNoFr").Visible = False: adc("UoLotNoFr") = ""
    adc("UoLotNoTo").Enabled = False: adc("UoLotNoTo").Visible = False: adc("UoLotNoTo") = ""
    adc("UoLotNoSel").Enabled = False: adc("UoLotNoSel").Visible = False: adc("UoLotNoSel") = ""
  End If
  ' **** Manali 350Nxt
  
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
    Case Is = UCase("UoInVal")     '4.1.4
        If UCase(adc("UoInVal")) <> UCase("S") Then adc("UoInVal2") = ""
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Or UCase(adc("UoInVal1")) = "S" Then Cancel = True: Exit Sub
    
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = "XS"
    
  Case Is = UCase("UoInVal2")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoInVal")) <> UCase("S") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
    
  End Select
  ' Zubin 213
  SetProp Me, IdName, When
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
  'StkRt-Set Show StockRate option is No
  adc("UoYN") = "N"

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
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

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  ' Zubin 213
  Select Case UCase(IdName)
  Case Is = UCase("UoCurCdFr")
    If UCase(adc("UoInVal")) = "S" Or UCase(adc("UoInVal1")) = "S" Then adc("UoCurCdFr") = "": Exit Sub
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
Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
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

Private Sub ChkBoxArr_Click(Index As Integer)
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
  With ChkBoxArr(Index)
  Select Case Index
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
  End Select
  End With
End Sub


'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(0).Value = Unchecked
    End If
End Sub
'StkRt-if user is Seo user then validate with their Seo password else checks with SEOPWD password
Private Sub TxtPwd_Validate(Cancel As Boolean)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select isnull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If

End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub

