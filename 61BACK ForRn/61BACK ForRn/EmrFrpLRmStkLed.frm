VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLRmStkLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Ledger (Loose Rm)"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   11010
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   54
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
      Left            =   4965
      TabIndex        =   63
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   64
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
         TabIndex        =   65
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
      Left            =   6180
      TabIndex        =   62
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   59
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
      TabIndex        =   55
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
      Height          =   9705
      Left            =   0
      TabIndex        =   58
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   57
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpLRmStkLed.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpLRmStkLed.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   15105
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1560
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   46
               Top             =   6165
               Width           =   2340
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Value"
               BeginProperty Font 
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
               Left            =   30
               TabIndex        =   47
               Tag             =   "RevLogic"
               ToolTipText     =   "Check To See Stock Value by Weight"
               Top             =   6450
               Width           =   1710
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Bal"
               BeginProperty Font 
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
               TabIndex        =   49
               Tag             =   "ShowDesc"
               ToolTipText     =   "Check To Show Openning/Closing Balance"
               Top             =   7020
               Width           =   1710
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Reverse Logic"
               BeginProperty Font 
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
               TabIndex        =   48
               Tag             =   "RevLogic"
               ToolTipText     =   "Check To Select Reverse Logic"
               Top             =   6735
               Width           =   1710
            End
            Begin VB.Frame FraOptQW 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1515
               TabIndex        =   86
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   5880
               Width           =   1980
               Begin VB.OptionButton OptQW 
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
                  Index           =   0
                  Left            =   840
                  TabIndex        =   45
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   825
               End
               Begin VB.OptionButton OptQW 
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
                  Index           =   1
                  Left            =   0
                  TabIndex        =   44
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   765
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3915
               TabIndex        =   8
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1575
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "LocCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6300
               TabIndex        =   9
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1575
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "LocCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   3915
               TabIndex        =   41
               ToolTipText     =   "Enter To Customer (Customer As Mentioned In Rm Master)"
               Top             =   5610
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "RmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6300
               TabIndex        =   42
               ToolTipText     =   "Enter Customer Selection (Customer As Mentioned In Rm Master)"
               Top             =   5610
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "RmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1515
               TabIndex        =   40
               ToolTipText     =   "Enter From Customer (Customer As Mentioned In Rm Master)"
               Top             =   5595
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "RmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   3915
               TabIndex        =   38
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   5040
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               ReCalcParent    =   "UOYN5"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   37
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   5025
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               ReCalcParent    =   "UOYN5"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1575
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "LocCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3915
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3015
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   3915
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2445
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
               Index           =   6
               Left            =   1515
               TabIndex        =   10
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1860
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
               Index           =   7
               Left            =   1515
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2430
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
               Index           =   9
               Left            =   1515
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3000
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   3915
               TabIndex        =   30
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   3870
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1515
               TabIndex        =   29
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   3855
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   3915
               TabIndex        =   36
               ToolTipText     =   "Enter To Voucher Rate"
               Top             =   4755
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "OdJLVchRt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1515
               TabIndex        =   35
               ToolTipText     =   "Enter From Voucher Rate"
               Top             =   4740
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "OdJLVchRt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1515
               TabIndex        =   3
               ToolTipText     =   "Enter Location Type"
               Top             =   990
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
               CmpStr          =   "LocTyp = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   3915
               TabIndex        =   27
               ToolTipText     =   "Enter To Raw Material Range Size"
               Top             =   3585
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
               CmpStr          =   "RrRngSz <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   26
               ToolTipText     =   "Enter From Raw Material Range Size"
               Top             =   3570
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZFR"
               IdName          =   "UORNGSZFR"
               CmpStr          =   "RrRngSz >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   3915
               TabIndex        =   18
               ToolTipText     =   "Enter To Parent Code"
               Top             =   2730
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDTO"
               IdName          =   "UOPARENTCDTO"
               CmpStr          =   "RmParentCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   1515
               TabIndex        =   17
               ToolTipText     =   "Enter From Parent Code"
               Top             =   2715
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOPARENTCDFR"
               IdName          =   "UOPARENTCDFR"
               CmpStr          =   "RmParentCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "LocCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   39
               ToolTipText     =   "Enter Open YearMonth"
               Top             =   5310
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOOPNYYMM"
               IdName          =   "UOOPNYYMM"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   3600
               TabIndex        =   43
               ToolTipText     =   "Enter 'N' for Quantity and 'Y' for Weight"
               Top             =   5880
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
               Index           =   24
               Left            =   2115
               TabIndex        =   52
               ToolTipText     =   "Enter 'Y' to apply Reverse Logic else enter 'N' for normal logic"
               Top             =   6735
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
               Index           =   32
               Left            =   1515
               TabIndex        =   23
               ToolTipText     =   "Enter From Lot No."
               Top             =   3285
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   3915
               TabIndex        =   24
               ToolTipText     =   "Enter To Lot No."
               Top             =   3300
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   3915
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2160
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "sc.PSCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2145
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "sc.PSCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6300
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3000
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6300
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2445
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6300
               TabIndex        =   28
               ToolTipText     =   "Enter Raw Material Range Size Selection"
               Top             =   3585
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
               CmpStr          =   "RrRngSz In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   6300
               TabIndex        =   19
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   2730
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   6300
               TabIndex        =   25
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3300
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6300
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2160
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "sc.PSCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2160
               TabIndex        =   51
               ToolTipText     =   "Enter 'Y' to Show  Value "
               Top             =   6450
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN6"
               IdName          =   "UOYN6"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   3915
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   690
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "LocCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6330
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   690
               Width           =   8640
               _ExtentX        =   15240
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "LocCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   3915
               TabIndex        =   5
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1290
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   6300
               TabIndex        =   6
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1290
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1515
               TabIndex        =   4
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1290
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   2115
               TabIndex        =   87
               ToolTipText     =   "Show Description (Yes/No)"
               Top             =   7020
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               ReCalcOn        =   "UODTFR,UODTTO"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   3915
               TabIndex        =   32
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   4155
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN2TO"
               IdName          =   "UOLN2TO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   1515
               TabIndex        =   31
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   4140
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN2FR"
               IdName          =   "UOLN2FR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   3915
               TabIndex        =   34
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   4440
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN3TO"
               IdName          =   "UOLN3TO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1515
               TabIndex        =   33
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   4425
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN3FR"
               IdName          =   "UOLN3FR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1530
               TabIndex        =   50
               ToolTipText     =   "Enter Currency Code Whose Vouchers You Want To View Or Leave It Blank To View All Vouchers In Base Currency"
               Top             =   7350
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
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
               TabIndex        =   92
               ToolTipText     =   "Location"
               Top             =   7350
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Value will be shown only if Report is taken By weight"
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
               Left            =   3555
               TabIndex        =   91
               Top             =   6480
               Width           =   5265
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "SEO Pwd"
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
               TabIndex        =   90
               Top             =   6165
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Sz 3"
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
               TabIndex        =   89
               Top             =   4425
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Sz 2"
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
               TabIndex        =   88
               Top             =   4140
               Width           =   1425
            End
            Begin VB.Label LblPrntCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Loc"
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
               TabIndex        =   85
               Top             =   1290
               Width           =   1425
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
               Index           =   18
               Left            =   60
               TabIndex        =   84
               Top             =   2145
               Width           =   1425
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
               Left            =   60
               TabIndex        =   83
               Top             =   3285
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "OpenYyyyMm"
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
               Left            =   60
               TabIndex        =   82
               Top             =   5310
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Wt/Qty"
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
               TabIndex        =   81
               Top             =   5880
               Width           =   1455
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
               Index           =   25
               Left            =   60
               TabIndex        =   80
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Parent Cd"
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
               Left            =   60
               TabIndex        =   79
               Top             =   2715
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Range Size"
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
               TabIndex        =   78
               Top             =   3570
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Location Typ"
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
               TabIndex        =   77
               Top             =   990
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Size/Ptr"
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
               TabIndex        =   76
               Top             =   3855
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
               Index           =   6
               Left            =   60
               TabIndex        =   75
               ToolTipText     =   "Location"
               Top             =   1860
               Width           =   1425
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
               Index           =   7
               Left            =   60
               TabIndex        =   74
               Top             =   2430
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
               Left            =   60
               TabIndex        =   73
               Top             =   3000
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Voucher Rate"
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
               TabIndex        =   72
               Top             =   4740
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Voucher Date"
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
               TabIndex        =   71
               Top             =   5025
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Rm Customer"
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
               TabIndex        =   70
               Top             =   5595
               Width           =   1425
            End
            Begin VB.Label LblLocCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
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
               TabIndex        =   69
               Top             =   1575
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
               Left            =   6300
               TabIndex        =   68
               Top             =   120
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
               Left            =   3915
               TabIndex        =   67
               Top             =   120
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
               TabIndex        =   66
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   61
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   56
               Top             =   360
               Width           =   10245
               _ExtentX        =   18071
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpLRmStkLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '
Dim Rep As New EmrRepLRmStkLed
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06
Dim ms_PwdFlag As Boolean
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
  GRP_REP.Add "Process", "LocCoCd+LocCd", "'['+LocCoCd+'] '+LocCd", "qLoc", "hqLoc", "Company Code,Parent Code", "", "", "LocDesc"
  GRP_REP.Add "Rm Ctg", "Rtrim(RmCtg)", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmCtg)+Rtrim(RmSCtg)", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
  GRP_REP.Add "Rm Code", "Rtrim(RmCtg)+Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ RmCd", "RmCtg,RmSCtg,RmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg,BaseCd + RmCd,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "BaseCd + RmCd", "Rtrim(RmBaseCd)+Rtrim(RmCd)", "'('+ RTrim(RmBaseCd) + ') '+ RmCd", "RmCd", "hRmCd", "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd", "", "", "SubString(RmDesc, 1, 30)"
  GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
'  GRP_REP.Add "Customer Code", Array(0, "OmCmCd"), "", "", "", "", "", "", Array("CmCd", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + gs_UsrCd + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCd= qCmCd)")
  GRP_REP.Add "Customer Code", "RTrim(RmCmCd)", "", "", "", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= RmCmCd) "
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
  
  'GRP_REP.Add "Rm Sv/Len", Array("case when " + wsPtrYN + "='Y' then str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)  else IsNull(RisRmSzDesc, str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2) else IsNull(RisRmSzDesc, str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2)) end "), _
  "", "qRmSz1,qRmSz2,qRmSz3", "hqRmSz1,hqRmSz2,hqRmSz3", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  
  GRP_REP.Add "Rm Sv/Len", Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else RisRmSzDesc end ) end ", _
  "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end ", _
  "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end "), _
  "", "wRmSz", "hqRmSz1", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  
  
  GRP_REP.Add "Rm VchRt", Array(0, "str(OdJlVchRt,9,2)"), "", "qVchRt", "hqVchRt", "(RmCd) VchRt", "", "", ""
    
  If (moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and " + _
                                                  "HCd='" + ctSelfCmCd + "'") = "Y") Or (ws_MultiCoMod = True) Then
      'GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)  else IsNull(RisRmSzDesc, str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2) else IsNull(RisRmSzDesc, str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2)) end "), _
      "", "qRmSz1,qRmSz2,qRmSz3", "hqRmSz1,hqRmSz2,hqRmSz3", "(RmCd) Sv/Len,Range Size", "", "", ""
      
      GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else RisRmSzDesc end ) end ", _
      "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end "), _
      "", "wRmSz", "hqRmSz1", "(RmCd) Sv/Len,Range Size", "", "", ""
  End If
  
  GRP_REP.Add "Range Size", "RmCtg+RmSCtg+IsNull(RrRngSz, '')", "'('+ RTrim(RmCtg)+ ', '+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", ""
   
  'GRP_REP.Add "(RmCd) Sv/Len", Array("RmCd+ case when " + wsPtrYN + "='Y' then str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2) else IsNull(RisRmSzDesc, str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2) else IsNull(RisRmSzDesc, str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2)) end "), _
     Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then  str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2) else IsNull(RisRmSzDesc,  str(DJLLn1,5,2)+str(DJLLn2,5,2)+str(DJLLn3,5,2)) end )", _
     " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2) else IsNull(RisRmSzDesc, str(OdJLLn1,5,2)+str(OdJLLn2,5,2)+str(OdJLLn3,5,2)) end )"), _
     "RmCtg,RmSCtg,RmCd,qRmSz1,qRmSz2,qRmSz3", "hRmCtg,hRmSCtg,hRmCd,hqRmSz1,hqRmSz2,hqRmSz3", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,Rm Sv/Len,Rm Ptr,(RmCd) StkRt,Range Size", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  
  GRP_REP.Add "(RmCd) Sv/Len", Array("RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else RisRmSzDesc end) end ", _
    "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end "), _
     Array(" '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("DJLLn1", "DJLLn2", "DJLLn3") + " else RisRmSzDesc end)  end)", _
      " '('+RmCd+') '+LTrim(case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end )"), _
      "RmCtg,RmSCtg,RmCd,wRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz1", "Rm Ctg,Rm SubCtg,Rm Code,BaseCd + RmCd,Rm Sv/Len,Rm Ptr,(RmCd) StkRt,Range Size", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
      
     
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(Sc.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(sc.PSCd, '')", "RmCtg", "hRmCtg", "Rm Ctg", "", "", "sc.PDesc"
  If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
    GRP_REP.Add "RM Lot No.", Array("DJLLotno", "OdJLLotno"), "", "qLotNo", "hLotNo", "", "", "", Array("DJLLotno", "OdJLLotno")
  End If
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "LocCoCd", "LocCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=LocCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= '" + ctSelfCoCd + "' And LocCd=Lc.LocPrntCd) "
  End If
  
End Sub

Private Sub adc_setreprecsource()
'On Error GoTo errhld
Dim ws_Val As String

  '*** Report Sql ***
  '*** (Jen 2.12)
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(1) as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz, RrRngSz, RmCtg, RmSCtg,
  '  OdJLVchRt as qStkRt, PValue3 as qPtrYN, OdJLWt as qPureWt, OdCoCd, OdTc, OdYy,
  '  OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt as qDt,
  '  LocPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, OdJlQty as qDrQty,
  '  OdJlQty as qCrQty, 0.0 as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt,
  '  OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt,
  '  space(1) As qQW, space(1) As qQWCst, OdJLVchRt as qVal
  '  From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2
  '*** (Jen 2.12)

  '*** (Bef 2.12)
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(1) as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz, RrRngSz, RmCtg, RmSCtg,
  '  OdJLVchRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, OdCoCd, OdTc, OdYy,
  '  OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt as qDt,
  '  LocPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, OdJlQty as qDrQty,
  '  OdJlQty as qCrQty, 0.0 as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt,
  '  OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt,
  '  space(1) As qQW
  '  From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2
  '*** (Bef 2.12)
  '*** Report Sql ***
  
  Dim grpflds0 As String, grpflds1 As String, grpflds2 As String, grpflds3 As String, wrepcnd As String
  Dim wCnd As String, wSqlStrg As String, wPurityWt As String, wDtCndOpnYyMm As String
  Dim wsqlOpnYyMm As String, wFrDt As String, wDtCndOpnTd As String, wSqlOpnTdFr As String
  Dim wSqlOpnTdTo As String, wDtCndDet As String, wSqlDetFr As String, wSqlDetTo As String
  Dim wSqlOpnFgFr As String, wSqlOpnFgTo As String, wSqlDetFgFr As String, wSqlDetFgTo As String
  Dim wFrDt1 As String, wDtCndOpnFg As String, wDtCndDetFg As String, wFgStr As String
  Dim wTdRate As String, wFrRate, wSyRate As String
  Dim wQw As String, wOpnYyStr As String
  Dim wTmpDt As String
  Dim wPureWtOrZWt As String, wXPureWtOrZwt As String
  
  '*****sachin 25-02-05
  Dim wSetTmpTab As String, wQwChk As String
  Dim wb_ShowOpn, wb_MonEnd, wb_MonStart As Boolean
  
  '*********** urmi 26/8/03 ************
  Dim wRLQty As String, wRLWt As String
  '*************************************
  '********************** urmi 5/7/03 ch28********************************************'
  Dim wDtCndOpnInv As String, wDtCndDetInv As String
  Dim wSqlOpnInvFr As String, wSqlOpnInvTo As String, wSqlDetInvFr As String, wSqlDetInvTo As String
  Dim wInvStr As String, wIrRate As String
  '************************************************************************************'
  Dim wHidLocStr As String
  '*** Jay 2.14 ***
  Dim qDesc As String
  '*** Jay 2.14 ***
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  Dim wRmSzDesc As String
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
    
  Dim wAddlCnd As String   '****** Sachin 3.02.0
  
  Dim wDtFr As String, wDtTo As String  ' ***** Manali 3.4.1 - 02/12/08
  
  ' ***** Manali Trading Module
  Dim wBagJoin As String
  Dim wSqlOpnXdFr As String, wSqlOpnXdTo As String, wSqlDetXdFr As String, wSqlDetXdTo As String
  Dim wDtCndOpnXd As String, wDtCndDetXd As String
  ' ***** Manali Trading Module
  
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******]
  ' **** Manali 3.4.1 - 02/12/08 - "01/01/80" excluded from check
  Dim ws_PurgeDt As String
  'sv.9 getting PurgeDate from vParam for ZZZ company code
  ws_PurgeDt = moCn.GetFldVal("Select vPDesc225 from vParam where vPCoCd ='ZZZ' and " + _
                                " vPTyp= 'JPURGEDT' and vPMCd= 'JPURGEDT' and vPSCd =''")
    
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And _
      ((CDate(adc("UOOPNYYMM")) <= Format(CDate(ws_PurgeDt), "YYYYMM") And adc("UOOPNYYMM") <> 0) Or _
        (CDate(adc("UODTFR")) <= CDate(ws_PurgeDt) And CDate(adc("UoDtFr")) <> MWLib.EmptyDate) Or _
        (CDate(adc("UODTTO")) <= CDate(ws_PurgeDt) And CDate(adc("UoDtTo")) <> MWLib.EmptyDate)) Then
      DispMsg "From/To Date Or OpenYYYYMM Cannot Be Less Than ZZZ Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError: Exit Sub
  End If
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  grpflds2 = GRP_REP.GrpFldLst(2)
  grpflds3 = GRP_REP.GrpFldLst(3)
  
  '*** Jay 2.14 ***
  '*** To check if Detail is selected ***
  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 3
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  '*** If Detail not selected then suppress Section 'DetDesc' Of Detail ***
  If wDetPos = 0 Then Rep.DETDesc.Suppress = True: Rep.DetSecLine.Suppress = True
  '*** Jay 2.14 ***
  
  ' ****** Manali 3.4.1 - "01/01/80" converted to purge dt
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And CDate(adc("UoDtFr")) = MWLib.EmptyDate Then
    wDtFr = CDate(ws_PurgeDt)
  Else
    wDtFr = adc("UoDtFr")
  End If
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And CDate(adc("UoDtTo")) = MWLib.EmptyDate Then
    wDtTo = CDate(ws_PurgeDt)
  Else
    wDtTo = adc("UoDtTo")
  End If
  
  Dim wMonEndDt As String
  wb_MonEnd = False: wb_MonStart = False
   If moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(wDtTo) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(wDtTo) + "')as Varchar(4))))," + CStr(DtFmtSlahVal()) + ")='" + wDtTo + "'") Then wb_MonEnd = True
   If CInt(Format(CStr(adc("UODTFr")), "dd")) = 1 Then wb_MonStart = True
    wb_ShowOpn = IIF(adc("UOYN5") = "Y", True, False)
    
    adc("UOYN2") = IIF(CInt(Format(wDtFr, "dd")) <= 15, "N", "Y")
    If wb_MonEnd Then adc("UOYN2") = "Y"

  If adc("UOYN2") = "Y" Then
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    '6.1 added in if and else
    If CInt(Format(wDtFr, "mm")) > 0 And CInt(Format(wDtFr, "mm")) < 12 Then
        wTmpDt = Format$(DateAdd("d", -1, CDate(IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format$(CInt(Format(wDtFr, "mm")) + 1, "00") + "/01", "01/" + Format$(CInt(Format(wDtFr, "mm")) + 1, "00")) + "/" + CStr(CInt(Format(adc("UoDtFr"), "yy"))))), DtShortStr())
    ElseIf CInt(Format(wDtFr, "mm")) = 12 Then
        wTmpDt = Format$(DateAdd("d", -1, CDate("01/01/" + CStr(CInt(Format(wDtFr, "yy")) + 1))), DtShortStr())
    End If
    
  End If
  
  If adc("UOYN3") = "Y" Then
    Rep.TxtHead.SetText ("Stock Ledger (Loose Rm) (By Quantity)")
    'Rep.RecordSelectionFormula = "{rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0"
  ElseIf adc("UOYN3") = "N" Then
    Rep.TxtHead.SetText ("Stock Ledger (Loose Rm) (By Weight)")
    'Rep.RecordSelectionFormula = "Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 Or Round({rdo.qOpnWt}, 3) <> 0"
  End If
  
'  If UCase(adc("UoYN1")) = UCase("Y") Then
'    Rep.h1PureWt.SetText ("Z Pure Wt")
'    Rep.h2PureWt.SetText ("Z Pure Wt")
'    wPureWtOrZWt = "RmPurityZ"
'    wXPureWtOrZwt = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=Rm.RmBaseCd " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=Rm.RmPrtKey)", ")")
'  ElseIf UCase(adc("UoYN1")) = UCase("N") Then
    wPureWtOrZWt = "RmPurityWt"
    wXPureWtOrZwt = "RmBasePurityWt"
'  End If

  
    wSyRate = "0 as qRmRt"
    wTdRate = "OdJLVchRt as qRmRt"
    wFrRate = "0 as qRmRt"
    wIrRate = "0 as qRmRt"
    
    'ws_Val = ", 0 as qVal "         '*** (Jen 2.12 Patch 1)
    ws_Val = ", sum(convert (decimal (16, 3), (qOpnWt+ qDrWt- qCrWt)* qStkRt )) as qVal "
    
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qVal}"
'    GRP_REP.SetFormula Rep, "wRmClsVal", IIF(wb_MonEnd, "if {rdo.qFlag}='O' Then {rdo.qVal} Else 0", "{rdo.qVal}")
    GRP_REP.SetFormula Rep, "wQW", "''"
    Rep.HDtRmVal.SetText "StkVal"
    Rep.HGrpRmVal.SetText "StkVal"
    
        
    '*** (Jen 2.13)
    If adc("UOYN6") = "N" Or adc("UOYN3") = "Y" Then
    '*** (Jen 2.13)
      Rep.wRmVal.Suppress = True
      Rep.g3RmVal.Suppress = True
      Rep.g2RmVal.Suppress = True
      Rep.g1RmVal.Suppress = True
      Rep.gRmVal.Suppress = True
      Rep.HDtRmVal.Suppress = True
      Rep.HGrpRmVal.Suppress = True
    End If
    '*** (Jen 2.13)
  '********************************'
'  End If
            
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
 '   wPurityWt = " /(Case When Rm.RmCtg In ('G', 'P', 'S', 'L', 'M') Then " + _
                          "IsNull((Select RmPurityWt from Param Join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                 "When Rm.RmCtg= 'X' then " + _
                          "IsNull((Select RmPurityWt from Param Join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                 "Else 1 end) "
  
  wPurityWt = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                 "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                 "else 1 end) "
 
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
'    wQwChk = " IsNull((Case when (RmCtg in ('G', 'P', 'S', 'L', 'M') And RmQw='W') Then (Select RmPr.RmQw From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue From param " + _
              "       Where PTyp= 'RMCTG' and PMCd= Rm.RmCtg)) Else RmQw End),'')  as qQw," + _
              "IsNull((Case when (RmCtg In ('G', 'P', 'S', 'L', 'M') And RmQwCst = 'W') Then (Select RmPr.RmQwCst From RmMst RmPr " + _
              "   Where " + IIF(gs_Partition = ctCurrPrtn, " RmPr.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmPr.RmCd=(Select PValue From param " + _
              "       Where PTyp= 'RMCTG' and PMCd = Rm.RmCtg)) Else RmQwCst End),'') as qQwCst "
              
      wQwChk = "'W' as qQwCst"
  ' ****** Sachin 2.13.0 - 05/06/2006 ******
    
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(1) as qFlag, RmCd, OdJLLn1 as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg,
'  OdJLVchRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, OdCoCd, OdTc, OdYy,
'  OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt as qDt,
'  LocPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, OdJlQty as qDrQty,
'  OdJlQty as qCrQty, 0.0 as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt,
'  OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt,
'  space(1) As qQW, space(1) As qQWCst
'  From Txnd, Txn, RmMst, Loc, RmRt, Param, RmIdSz where 1=2

'  "Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
'  "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
'  "space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
'  "space(1) as qFlag, RmCd, OdJLLn1 as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg, " + _
'  "OdJLVchRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, OdCoCd, OdTc, OdYy, " + _
'  "OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt as qDt, " + _
'  "LocPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, OdJlQty as qDrQty, " + _
'  "OdJlQty as qCrQty, 0.0 as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt, " + _
'  "OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt, " + _
'  "space(1) As qQW, space(1) As qQWCst " + _
'  "Into #TmpLRmStkLed From Txnd, Txn, RmMst, Loc, RmRt, Param where 1=2 " + _

  'Setting the opening tmp table
  'wSetTmpTab = "Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
            "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,  " + _
            "space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
            "Space(1) as qFlag, RMCd, DJLLn1, RisRmSzDesc as qRmSzDesc, IsNull(RrRngSz, '') RrRngSz, RmCtg, " + _
            "RmSCtg, SRmStkRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, " + _
            "OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt, " + _
            "LocCd, (SRmDrQty - SRmCrQty) as qOpnQty, SRmDrQty as qDrQty, " + _
            "SRmCrQty as qCrQty, (DDmDrWt- DDmCrWt) as qOpnWt, DDmDrWt as qDrWt, DDmCrWt as qCrWt, " + _
            "OmFrDmLoc as qFrRmLoc, OmToDmLoc as qToRmLoc, OdJLVchRt as qRmRt, RmQW as qQw, RmQWCst as qQwCst  " + _
            "Into #TmpLRmStkLed From DYyMm, Txn, Txnd, Loc, RmMst, RmRt, Param Where 1=2"
  ' Zubin 212
  
'*** Jay 2.14 *** (TDesc as qDesc added)
' *** Manali 3.03 - (RisRmSzDesc as qRmSzDesc) added
'*** REPORT SQL ***
'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'space(1) as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, PSCd As qStnCls, RmSCtg,
'OdJLVchRt as qStkRt, PValue3 as qPtrYN, OdJLWt as qPureWt, OdCoCd, OdTc, OdYy,
'OdChr, OdNo, OdSr, OdSrNo, OdDmStkYy, OdDmStkChr, OdDmStkNo, TDt as qDt, TDesc as qDesc,
'LocPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, OdJlQty as qDrQty,
'OdJlQty as qCrQty, 0.0 as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt,
'OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt,
'space(1) As qQW, space(1) As qQWCst, OdJLVchRt as qVal
'From Txnd, Txn, RmMst, Loc, RmRt, Param, RmIdSz where 1=2
'*** Jay 2.14 ***

'*** Jay 2.14 ***
'If UCase(adc("UoYn4")) = "Y" Then
'  If UCase(adc("UoVchOpt")) = UCase("Vch") Then qDesc = " TDesc " Else qDesc = " TdDesc "
'Else
  qDesc = " '' "
  Rep.DETDesc.Suppress = True
'End If
'*** Jay 2.14 ***
'*** Jay 2.14 ***(Tdesc as qDesc, added)
' **** Manali 3.03 - 09/07/08 - 'RisRmSzDesc as qRmSzDesc' addded
  wSetTmpTab = "Select space(40) as Grp1, space(40) as Grp2, space(40) as Grp3, " + _
            " space(40) as DGrp1, space(40) as DGrp2, space(40) as DGrp3, " + _
            " space(40) as G1Desc, space(40) as G2Desc, space(40) as G3Desc, " + _
            " space(1) as qFlag, RmCd, DJLLotno as qLotNo, DJLLn1 as qRmSz1, DJLLn2 as qRmSz2, DJLLn3 as qRmSz3, RisRmSzDesc as qRmSzDesc, RrRngSz, RmCtg, " + _
            " PSCd As qStnCls, RmSCtg, " + _
            " OdJLVchRt as qStkRt, PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, " + _
            " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt as qDt, OmCmCd as qCmCd, '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo, '' as qDesc, " + _
            " LocPrntCd as qPrntCd, LocCd as qLoc,OdJLQty as qOpnQty, OdJLQty as qDrQty, OdJLQty as qCrQty, OdJLWt as qOpnWt, OdJLWt as qDrWt, OdJLWt as qCrWt, " + _
            " OmFrDmLoc As qFrRmLoc, OmToDmLoc As qToRmLoc, OdJLVchRt As qRmRt, 'W' As qQw, 'W' as qQwCst, PValue1 as qStkbyRt " + _
            " Into #TmpLRmStkLed From DmYyMm, OrdMst, OrdDsg, Loc, RmMst, RmRt, Param, RmIdSz where 1=2 "

  moCn.CreateTmpTable "#TmpLRmStkLed", wSetTmpTab
  
  adc("UoLn1Fr").CmpStr = "DJLLn1>= "
  adc("UoLn1To").CmpStr = "DJLLn1<= "
  adc("UoLn2Fr").CmpStr = "DJLLn2>= "
  adc("UoLn2To").CmpStr = "DJLLn2<= "
  adc("UoLn3Fr").CmpStr = "DJLLn3>= "
  adc("UoLn3To").CmpStr = "DJLLn3<= "
  adc("UoCstRtFr").CmpStr = ""
  adc("UoCstRtTo").CmpStr = ""
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "DJLLotno>= "
  adc("UoLotNoTo").CmpStr = "DJLLotno<= "
  adc("UoLotNoSel").CmpStr = "DJLLotno In "
  
  wrepcnd = adc.RepCond
  
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd , "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd + " Option (MaxDop 1)", " Option (MaxDop 1)")
  
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnYyMm = " where DYyMm  " + IIF(wb_MonEnd, "<=", "<") + Format(IIF(wb_MonEnd, wDtTo, wDtFr), "yyyymm")
  'wDtCndOpnYyMm = " where DYyMm < " + Format(adc("UoDtFr"), "yyyymm")
  
  If adc("UoOpnYyMm") <> 0 Then
    wOpnYyStr = " and DYyMm >= " + CStr(adc("UoOpnYyMM")) + ""
  Else
    wOpnYyStr = ""
  End If
  
  '*****************  Reverse Logic urmi 2.04 *******************
  ' eg. UoDtFr = 20/7/2002 UoDtTo 22/4/03 and Todays dt = 25/8/03
  ' wDtCndOpnYyMm = "Where DYyMm <= 200207 "
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  If adc("UoYn2") = "Y" Then
    If wb_MonEnd Then
        wDtCndOpnYyMm = " where DYyMm <= " + Format(wDtTo, "yyyymm")
    Else
        wDtCndOpnYyMm = " where DYyMm <= " + Format(wDtFr, "yyyymm")
    End If
  End If
  'If adc("UoYn2") = "Y" Then wDtCndOpnYyMm = " where DYyMm <= " + Format(adc("UoDtFr"), "yyyymm")
  '**************************************************************
  
  
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End "
  '*** Manali 3.03 - 07/07/08 - RmSzDesc
  '578.22 getting base currency
  Dim ws_BaseCurCd As String, wb_BaseCur As Boolean, wCurCdCnd As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  '578.22 if currency code not given it takes base currency from head else takes from scope
  If adc("UoCurCdFr") = "" Then
    wb_BaseCur = True
    wCurCdCnd = ""
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_BaseCur = False
    wCurCdCnd = " and OmCmCurCd= '" + adc("UoCurCdFr") + "' "
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If
    
  '"RmPurityWt" + wPurityWt + " as qRmPurityWt, "
  '1
  ' Zubin 212
  ' ****** Sachin 2.13.0 - StnCls added to all queries ******
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  wsqlOpnYyMm = "Insert Into #TmpLRmStkLed Select " + grpflds0 + ", " + _
                "'O' as qFlag, RmCd, DJLLotno as qLotNo, DJLLn1 as qRmSz1, DJLLn2 as qRmSz2, DJLLn3 as qRmSz3," + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                "0 as qStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "DCoCd as OdCoCd, '' as OdTc, '' as OdYy, '' as OdChr, 0 as OdNo, 0 as OdSr, '' as qDt, '' as qCmCd,  '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo, '' as qDesc," + _
                "LocPrntCd as qPrntCd, DLoc as qLoc, " + _
                "(DDmDrQty- DDmCrQty) as qOpnQty, 0 as qDrQty, 0 as qCrQty, " + _
                "(DDmDrWt- DDmCrWt) as qOpnWt, 0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, 0 as qRmRt, " + _
                "'W' As qQw, 'W' as qQwCst,  RSCtg.PValue1 as qStkbyRt " + _
                "From DmYyMm NoLock Join Loc Lc On LocCoCd = DCoCd and LocCd= DLoc and LocHideYN='N'  " + _
                " Join RmMst Rm On RmCd= DJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=DPrtKey ", "") + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg " + _
                " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
                " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= DJLLn1 and RrToLn>= DJLLn1 " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=DJLLn1 " + _
                wDtCndOpnYyMm + wOpnYyStr + IIF(gs_Partition = ctCurrPrtn, " and DPrtKey='" + ctCurrPrtn + "' ", " and DYyMm<>'0' ") + wCnd
                
  If wb_ShowOpn Then moCn.Execute wsqlOpnYyMm
  
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wFrDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wDtFr, "mm") + "/01", "01/" + Format(wDtFr, "mm")) + "/" + Format(wDtFr, "yy") '6.1
  'wFrDt = "01/" + Format(adc("UoDtFr"), "mm") + "/" + Format(adc("UoDtFr"), "yy")
  '****** Sachin 3.02.0 - New fields
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndOpnTd = " where OmDt >= '" + wFrDt + "' And OmDt < '" + CStr(wDtFr) + "' "
  'wDtCndOpnTd = " where OmDt >= '" + wFrDt + "' And OmDt < '" + CStr(adc("UoDtFr")) + "' And TdDt >= '" + wFrDt + "' And TdDt < '" + CStr(adc("UoDtFr")) + "' "
  
  '************* Reverse Logic urmi 5/7/03 *************************************
  wRLQty = "(Case OmFrDmDc When 'D' Then OdJlQty When 'C' Then -OdJlQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case OmFrDmDc When 'D' Then OdJLWt When 'C' Then -OdJLWt Else 0 End) as qOpnWt, "
  
  If adc("Uoyn2") = "Y" Then
'    wDtCndOpnTd = " where OmDt >= '" & ADC("UoDtFr") & "' And month(TDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "mm")) & " and Year(TDt) = " & _
'                  CInt(Format(ADC("UoDtFr"), "yyyy"))

    '****** Sachin 3.02.0 - New fields
    ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
    wDtCndOpnTd = " where OmDt >= '" & wDtFr & "' And OmDt <= '" & wTmpDt & "' "
    'wDtCndOpnTd = " where OmDt >= '" & adc("UoDtFr") & "' And OmDt <= '" & wTmpDt & "' And TdDt >= '" & adc("UoDtFr") & "' And TdDt <= '" & wTmpDt & "' "
    
    wRLQty = "(Case OmFrDmDc When 'D' Then -OdJlQty When 'C' Then OdJlQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case OmFrDmDc When 'D' Then -OdJLWt When 'C' Then OdJLWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************************************
  adc("UoLn1Fr").CmpStr = "OdJLLn1>= "
  adc("UoLn1To").CmpStr = "OdJLLn1<= "
  adc("UoLn2Fr").CmpStr = "OdJLLn2>= "
  adc("UoLn2To").CmpStr = "OdJLLn2<= "
  adc("UoLn3Fr").CmpStr = "OdJLLn3>= "
  adc("UoLn3To").CmpStr = "OdJLLn3<= "
  adc("UoCstRtFr").CmpStr = "OdJLVchRt>= "
  adc("UoCstRtTo").CmpStr = "OdJLVchRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "OdJLLotno>= "
  adc("UoLotNoTo").CmpStr = "OdJLLotno<= "
  adc("UoLotNoSel").CmpStr = "OdJLLotno In "
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
     '****** Sachin 3.02.0 - Additional Conditions For New Fields
      wAddlCnd = ""
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " RmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
      
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " RmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " RmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " RmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields

  '2
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  '578.22 conversion factor added to OdJLVchRt, wCurCdCnd added
  wSqlOpnTdFr = "Insert Into #TmpLRmStkLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz1, OdJLLn2 as qRmSz2, OdJLLn3 as qRmSz3, " + _
                wRmSzDesc + " as qRmSzDesc,  IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OdJLVchRt/OmCnvFct else OdJLVchRt*OmCnvFct end) ", "OdJLVchRt") + " as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "OmCoCd as OdCoCd, '' as OdTc,'' as OdYy, '' as OdChr, 0 as OdNo, 0 as OdSr, " + _
                " '' as qDt, '' as qCmCd, '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo,'' as qDesc, LocPrntCd as qPrntCd, OmFrDmLoc as qLoc, " + _
                wRLQty + "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wTdRate + _
                ", 'W' As qQw, 'W' as qQwCst,  RSCtg.PValue1 as qStkbyRt " + _
                " From OrdMst NoLock " + _
                " Join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
                " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                " Join Loc Lc On OdCoCd= LocCoCd and OmFrDmLoc= LocCd and LocHideYN='N' " + _
                " Join RmMst Rm On OdJLRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OdPrtKey ", "") + _
                " Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=RmSCtg  " + _
                " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
                " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= OdJLLn1 and RrToLn>= OdJLLn1 " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OdJLLn1 " + _
                 wDtCndOpnTd + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " And LocTyp= '" + adc("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnTdFr
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '**************urmi 2.04  Reverse Logic ****************
  wRLQty = "(Case OmToDmDc When 'D' Then OdJlQty When 'C' Then -OdJlQty Else 0 End) as qOpnQty, "
  wRLWt = "(Case OmToDmDc When 'D' Then OdJLWt When 'C' Then -OdJLWt Else 0 End) as qOpnWt, "
  
  If adc("UoYn2") = "Y" Then
    wRLQty = "(Case OmToDmDc When 'D' Then -OdJlQty When 'C' Then OdJlQty Else 0 End) as qOpnQty, "
    wRLWt = "(Case OmToDmDc When 'D' Then -OdJLWt When 'C' Then OdJLWt Else 0 End) as qOpnWt, "
  End If
  '*******************************************************
  '3
  ' Zubin 212
  '*** jay 2.14 *** ('' as qDesc,)
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  '578.22 conversiom factor added to OdJLVchRt, wCurCdCnd added
  wSqlOpnTdTo = "Insert Into #TmpLRmStkLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz1, OdJLLn2 as qRmSz2, OdJLLn3 as qRmSz3, " + _
                wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
                IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OdJLVchRt/OmCnvFct else OdJLVchRt*OmCnvFct end) ", "OdJLVchRt") + " as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
                "OmCoCd as OdCoCd, '' as OdTc,'' as OdYy, '' as OdChr, 0 as OdNo, 0 as OdSr, " + _
                " '' as qDt, '' as qCmCd, '' as OdDmStkYy, '' as OdDmStkChr, 0 as OdDmStkNo,'' as qDesc, LocPrntCd as qPrntCd, OmToDmLoc as qLoc, " + _
                wRLQty + "0 as qDrQty, 0 as qCrQty, " + _
                wRLWt + "0 as qDrWt, 0 as qCrWt, '' as qFrRmLoc, '' as qToRmLoc, " + wTdRate + _
                ", 'W' As qQw, 'W' as qQwCst,  RSCtg.PValue1 as qStkbyRt " + _
                " From OrdMst NoLock " + _
                " Join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
                " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
                "   Join Loc Lc On OdCoCd= LocCoCd and OmToDmLoc= LocCd and LocHideYN='N' Join RmMst Rm On OdJLRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=RmPrtKey ", "") + _
                "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=RmSCtg  " + _
                "   Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
                "   Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "       RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "       RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= OdJLLn1 and RrToLn>= OdJLLn1 " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OdJLLn1 " + _
                wDtCndOpnTd + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + " And LocTyp= '" + adc("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  If wb_ShowOpn And Not wb_MonStart And Not wb_MonEnd Then moCn.Execute wSqlOpnTdTo
  
  '****** Sachin 3.02 - New fields
  ' ****** Manali 3.4.1 - 02/12/08 - [UoDtFr, UoDtTo respectively replaced by wDtFr, wDtTo]
  wDtCndDet = " where OmDt between '" + CStr(wDtFr) + "'" + _
              " And '" + CStr(wDtTo) + "'"
  'wDtCndDet = " where OmDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' And TdDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"
  
  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  adc("UoLotNoFr").CmpStr = "OdJLLotno>= "
  adc("UoLotNoTo").CmpStr = "OdJLLotno<= "
  adc("UoLotNoSel").CmpStr = "OdJLLotno In "
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '4
  ' Zubin 212
  '*** jay 2.14 *** (" + qDesc + " as qDesc, added)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  '4.1.4
  '578.22 conversion factor added to OdJLVchRt, wCurCdCnd added
  wSqlDetFr = " Insert Into #TmpLRmStkLed Select " + grpflds1 + "," + _
              " 'T' as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz1, OdJLLn2 as qRmSz2, OdJLLn3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc,  IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OdJLVchRt/OmCnvFct else OdJLVchRt*OmCnvFct end) ", "OdJLVchRt") + " as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt as qDt, OmCmCd as qCmCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, " + _
               qDesc + " as qDesc, LocPrntCd as qPrntCd, OmFrDmLoc as qLoc, 0 as qOpnQty, " + _
              " Case OmFrDmDc When 'D' Then OdJlQty Else 0 End  as qDrQty, Case OmFrDmDc When 'C' Then OdJlQty else 0 End  as qCrQty, " + _
              " 0 as qOpnWt, Case OmFrDmDc When 'D' Then OdJLWt Else 0 End  as qDrWt, " + _
              " Case OmFrDmDc When 'C' Then OdJLWt else 0 End  as qCrWt, " + _
              " (case when OmFrDmDc= 'C' then OmFrDmLoc else OmToDmLoc end) as qFrRmLoc, " + _
              " (case when OmFrDmDc= 'D' then OmFrDmLoc else OmToDmLoc end) as qToRmLoc, " + wTdRate + _
              " , 'W' As qQw, 'W' as qQwCst,  RSCtg.PValue1 as qStkbyRt " + _
              " From OrdMst NoLock " + _
              " Join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
               " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= OdCoCd and LocCd= OmFrDmLoc and LocHideYN='N' " + _
              " Join RmMst Rm On OdJLRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OdPrtKey ", "") + _
              "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              "   Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
              "   Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and  RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= OdJLLn1 and RrToLn>= OdJLLn1 " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OdJLLn1 " + _
              wDtCndDet + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " And LocTyp= '" + adc("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetFr

  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212 (This CmpStr block is not req as it is the same for the above and below query)
  adc("UoLotNoFr").CmpStr = "OdJLLotno>= "
  adc("UoLotNoTo").CmpStr = "OdJLLotno<= "
  adc("UoLotNoSel").CmpStr = "OdJLLotno In "
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  '5
  ' Zubin 212
  ' *** jay 2.14 ***(" + qDesc + " as qDesc, added)
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmIdSz added
  ' 4.1.4
  '578.22 conversion factor added to OdJLVchRt, wCurCdCnd added
  wSqlDetTo = " Insert Into #TmpLRmStkLed Select " + grpflds1 + "," + _
              " 'T' as qFlag, RmCd, OdJLLotno as qLotNo, OdJLLn1 as qRmSz1, OdJLLn2 as qRmSz2, OdJLLn3 as qRmSz3, " + wRmSzDesc + " as qRmSzDesc, IsNull(RrRngSz, '') as RrRngSz, RmCtg, IsNull(SC.PSCd,'') As qStnCls, RmSCtg, " + _
              IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OdJLVchRt/OmCnvFct else OdJLVchRt*OmCnvFct end) ", "OdJLVchRt") + " as qStkRt, (Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
              " (Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 0 end)" + wPurityWt + " as qRmPurityWt, " + _
              " OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr,  OmDt as qDt, OmCmCd as qCmCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, " + _
              qDesc + " as qDesc, LocPrntCd as qPrntCd, OmToDmLoc as qLoc, 0 as qOpnQty, " + _
              " Case OmToDmDc When 'D' Then OdJlQty Else 0 End  as qDrQty, " + _
              " Case OmToDmDc When 'C' Then OdJlQty else 0 End  as qCrQty, " + _
              " 0 as qOpnWt, Case OmToDmDc When 'D' Then OdJLWt Else 0 End  as qDrWt, " + _
              " Case OmToDmDc When 'C' Then OdJLWt else 0 End  as qCrWt, " + _
              " (case when OmFrDmDc= 'C' then OmFrDmLoc else OmToDmLoc end) as qFrRmLoc, " + _
              " (case when OmFrDmDc= 'D' then OmFrDmLoc else OmToDmLoc end) as qToRmLoc, " + wTdRate + _
              " , 'W' As qQw, 'W' as qQwCst,  RSCtg.PValue1 as qStkbyRt " + _
                " From OrdMst NoLock " + _
                " Join OrdDsg on OdOmIdNo=OmIdNo and OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy " + _
                " and OdChr= OmChr and OdNo= OmNo " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OmPrtKey ", "") + _
              " Join Loc Lc On LocCoCd= OdCoCd and LocCd= OmToDmLoc and LocHideYN='N' Join RmMst Rm On OdJLRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=RmPrtKey ", "") + _
              "   Join Param RSCtg On RSCtg.PTyp='RMSCTG' And RSCtg.PMCd=Rm.RmCtg And RSCtg.PSCd=Rm.RmSCtg  " + _
              " Left Outer Join Param SC on SC.PTyp= 'STNCLS' and SC.PmCd= RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
              " Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              " RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              " RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= OdJLLn1 and RrToLn>= OdJLLn1 " + _
              " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OdJLLn1 " + _
              wDtCndDet + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " And LocTyp= '" + adc("UOLOCTYPFR") + "'" + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetTo

  Dim ws_DetGrp As String
'  If UCase(adc("UoVchOpt")) = UCase("VchSrNo") Then
'    ws_DetGrp = ",OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OdSrNo "
'  ElseIf UCase(adc("UoVchOpt")) = UCase("VchSr") Then
    ws_DetGrp = ",OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr "
'  ElseIf UCase(adc("UoVchOpt")) = UCase("Vch") Then
'    ws_DetGrp = ",OdCoCd, OdTc, OdYy, OdChr, OdNo "
'  End If
  
  '*** jay 2.14 *** (Dt: 28/12/06)
  Dim ws_Desc As String
  ws_Desc = " Max(qDesc) "
  '*** jay 2.14 *** (Dt: 28/12/06)
  
  
  '*** jay 2.14 *** (max(qDesc) as qDesc,)
  '*** jay 2.14 *** (Dt: 28/12/06)
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc added
  adc.RepRecSource = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
      "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
      "max(qFlag) as qFlag, max(RmCd) as RmCd, max(qLotNo) as qLotNo, " + _
      "max(qRmSz1) as qRmSz1, max(qRmSz2) as qRmSz2, max(qRmSz3) as qRmSz3, max(qRmSzDesc) as qRmSzDesc, max(RrRngSz) as RrRngSz, max(RmCtg) as RmCtg, max(qStnCls) as qStnCls, " + _
      "max(RmSCtg) as RmSCtg, max(qStkRt) as qStkRt, max(qPtrYN) as qPtrYN,Sum(qOpnWt*qRmPurityWt) as qOpnPureWt, " + _
      "sum((qDrWt- qCrWt)* qRmPurityWt) as qPureWt, OdCoCd, OdTc, OdYy, OdChr, OdNo, max(OdSr) as OdSr, " + _
      "max(OdDmStkYy) as OdDmStkYy, max(OdDmStkChr) as OdDmStkChr, " + _
      "max(OdDmStkNo) as OdDmStkNo, max(qDt) as qDt, max(qCmCd) as qCmCd, " + ws_Desc + " as qDesc, max(qPrntCd) as qPrntCd, max(qLoc) as qLoc, sum(qOpnQty) as qOpnQty, " + _
      "sum(qDrQty) as qDrQty, sum(qCrQty) as qCrQty, sum(qOpnWt) as qOpnWt, " + _
      "sum(qDrWt) as qDrWt, sum(qCrWt) as qCrWt, max(qFrRmLoc) as qFrRmLoc, " + _
      "max(qToRmLoc) as qToRmLoc, max(qRmRt) as qRmRt, max(qQW) as qQW, max(qQWCst) as qQWCst " + ws_Val + _
      ", '' as qRmSz " + _
      "From #TmpLRmStkLed Tmp Group By DGrp1, DGrp2, DGrp3 " + ws_DetGrp
             
  Rep.TxtDtFr.SetText adc("UoDtFr")
  Rep.TxtDtTo.SetText adc("UoDtTo")
             
  Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
  Rep.TxtLocTyp.SetText adc("UoLocTypFr")
  Rep.TxtLocFr.SetText adc("UoIssLocFr")
  Rep.TxtLocTo.SetText adc("UoIssLocTo")
  Rep.TxtLocSel.SetText adc("UoIssLocSel")
  
  If UCase(adc("UoYN3")) = UCase("Y") Then
    wQw = "Qty"
    
    Rep.Fldg3Opn.DecimalPlaces = 0: Rep.Fldg3Opn.RoundingType = crRoundToUnit
    Rep.Fldg2Opn.DecimalPlaces = 0: Rep.Fldg2Opn.RoundingType = crRoundToUnit
    Rep.Fldg1Opn.DecimalPlaces = 0: Rep.Fldg1Opn.RoundingType = crRoundToUnit
    Rep.FldgOpn.DecimalPlaces = 0: Rep.FldgOpn.RoundingType = crRoundToUnit
  
    Rep.Fldg4Cr.DecimalPlaces = 0: Rep.Fldg4Cr.RoundingType = crRoundToUnit
    Rep.Fldg3Cr.DecimalPlaces = 0: Rep.Fldg3Cr.RoundingType = crRoundToUnit
    Rep.Fldg2Cr.DecimalPlaces = 0: Rep.Fldg2Cr.RoundingType = crRoundToUnit
    Rep.Fldg1Cr.DecimalPlaces = 0: Rep.Fldg1Cr.RoundingType = crRoundToUnit
    Rep.FldgCr.DecimalPlaces = 0: Rep.FldgCr.RoundingType = crRoundToUnit
  
    Rep.Fldg4Dr.DecimalPlaces = 0: Rep.Fldg4Dr.RoundingType = crRoundToUnit
    Rep.Fldg3Dr.DecimalPlaces = 0: Rep.Fldg3Dr.RoundingType = crRoundToUnit
    Rep.Fldg2Dr.DecimalPlaces = 0: Rep.Fldg2Dr.RoundingType = crRoundToUnit
    Rep.Fldg1Dr.DecimalPlaces = 0: Rep.Fldg1Dr.RoundingType = crRoundToUnit
    Rep.FldgDr.DecimalPlaces = 0: Rep.FldgDr.RoundingType = crRoundToUnit
  
    Rep.Fldg3Cls.DecimalPlaces = 0: Rep.Fldg3Cls.RoundingType = crRoundToUnit
    Rep.Fldg2Cls.DecimalPlaces = 0: Rep.Fldg2Cls.RoundingType = crRoundToUnit
    Rep.Fldg1Cls.DecimalPlaces = 0: Rep.Fldg1Cls.RoundingType = crRoundToUnit
    Rep.FldgCls.DecimalPlaces = 0: Rep.FldgCls.RoundingType = crRoundToUnit
  ElseIf UCase(adc("UoYN3")) = UCase("N") Then
    wQw = "Wt"
  End If
  
  '*** (Bef 2.12)
  'GRP_REP.SetFormula Rep, "g4_CrWt", " Sum ({rdo.qCr" + wQw + "}, {@wVchNo})"
  'GRP_REP.SetFormula Rep, "g4_DrWt", " Sum ({rdo.qDr" + wQw + "}, {@wVchNo})"
  '*** (Bef 2.12)
  
  '*** (Jen 2.12)
  GRP_REP.SetFormula Rep, "g4_CrWt", "{rdo.qCr" + wQw + "}"
  GRP_REP.SetFormula Rep, "g4_DrWt", "{rdo.qDr" + wQw + "}"
  '*** (Jen 2.12)
  
 If wb_MonEnd Then
            GRP_REP.SetFormula Rep, "g3_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_ClsWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_ClsWt", "Sum ({rdo.qOpn" + wQw + "})"
            
          If adc("UOYN5") = "Y" Then
            GRP_REP.SetFormula Rep, "g3_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_PureWt", "Sum ({rdo.qOpnPureWt}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_PureWt", "Sum ({rdo.qOpnPureWt})"
          End If
  
  Else
            GRP_REP.SetFormula Rep, "g3_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_OpnWt", "Sum ({rdo.qOpn" + wQw + "}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_OpnWt", "Sum ({rdo.qOpn" + wQw + "})"
            
           If adc("UOYN5") = "Y" Then
            GRP_REP.SetFormula Rep, "g3_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp3})+Sum ({rdo.qOpnPureWt}, {rdo.Grp3})"
            GRP_REP.SetFormula Rep, "g2_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp2})+Sum ({rdo.qOpnPureWt}, {rdo.Grp2})"
            GRP_REP.SetFormula Rep, "g1_PureWt", "Sum ({rdo.qPureWt}, {rdo.Grp1})+Sum ({rdo.qOpnPureWt}, {rdo.Grp1})"
            GRP_REP.SetFormula Rep, "g_PureWt", "Sum ({rdo.qPureWt})+Sum ({rdo.qOpnPureWt})"
           End If
  End If
  
  
  GRP_REP.SetFormula Rep, "g3_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_CrWt", "Sum ({rdo.qCr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_CrWt", "Sum ({rdo.qCr" + wQw + "})"
  GRP_REP.SetFormula Rep, "g3_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_DrWt", "Sum ({rdo.qDr" + wQw + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_DrWt", "Sum ({rdo.qDr" + wQw + "})"
  
'  GRP_REP.SetFormula Rep, "wRmSz", "ToText({rdo.qRmSz}, 3)"
  
 If wb_ShowOpn Then
      If wb_MonEnd Then
            
            GRP_REP.SetFormula Rep, "g3_OpnWt", "{@g3_ClsWt}-{@g3_DrWt} + {@g3_CrWt}"
            GRP_REP.SetFormula Rep, "g2_OpnWt", "{@g2_ClsWt}- {@g2_DrWt}+ {@g2_CrWt}"
            GRP_REP.SetFormula Rep, "g1_OpnWt", "{@g1_ClsWt}- {@g1_DrWt}+ {@g1_CrWt}"
            GRP_REP.SetFormula Rep, "g_OpnWt", "{@g_ClsWt}- {@g_DrWt}+ {@g_CrWt}"
    
       Else
            GRP_REP.SetFormula Rep, "g3_ClsWt", "{@g3_OpnWt}+ {@g3_DrWt}- {@g3_CrWt}"
            GRP_REP.SetFormula Rep, "g2_ClsWt", "{@g2_OpnWt}+ {@g2_DrWt}- {@g2_CrWt}"
            GRP_REP.SetFormula Rep, "g1_ClsWt", "{@g1_OpnWt}+ {@g1_DrWt}- {@g1_CrWt}"
            GRP_REP.SetFormula Rep, "g_ClsWt", "{@g_OpnWt}+ {@g_DrWt}- {@g_CrWt}"
     End If
  Else
    GRP_REP.SetFormula Rep, "g3_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g2_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g1_ClsWt", "0"
  GRP_REP.SetFormula Rep, "g_ClsWt", "0"
  End If

  '*** Set The Formula For Group4 According To The User Selection (Group4 is the Detail Record)
'  If UCase(adc("UoVchOpt")) = UCase("VchSrNo") Then
    '*** (Bef 15/05/07)
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdCoCd})+'/'+({rdo.OdTc})+'/'+" + _
    '"Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))+'/'+" + _
    '"Trim (ToText ({rdo.OdSr}, '##0'))+'/'+Trim (ToText ({rdo.OdSrNo}, '#0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
'    GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdTc})+'/'+" + _
    "Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))+'/'+" + _
    "Trim (ToText ({rdo.OdSr}, '##0'))+'/'+Trim (ToText ({rdo.OdSrNo}, '#0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    
 ' ElseIf UCase(adc("UoVchOpt")) = UCase("VchSr") Then
  
    '*** (Bef 15/05/07)
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdCoCd})+'/'+({rdo.OdTc})+'/'+" + _
    '"Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))+'/'+" + _
    '"Trim (ToText ({rdo.OdSr}, '##0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
    GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdCoCd}) + '/'+ Trim ({rdo.OdTc})+'/'+" + _
    "Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))+'/'+" + _
    "Trim (ToText ({rdo.OdSr}, '##0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
      
'    Rep.hqLoc.Suppress = True
'    Rep.hRmCtg.Suppress = True
'    Rep.hRmSCtg.Suppress = True
'    Rep.hRmCd.Suppress = True
'    Rep.hqRmSz.Suppress = True
    'Rep.hRrRngSz.Suppress = True
'    Rep.hqStkRt.Suppress = True
'    Rep.hFrRmLoc.Suppress = True
'    Rep.hToRmLoc.Suppress = True
      
'    Rep.FldqLoc.Suppress = True
'    Rep.FldRmCtg.Suppress = True
'    Rep.FldRmSCtg.Suppress = True
'    Rep.FldRmCd.Suppress = True
'    Rep.FldqRmSz.Suppress = True
'    Rep.FldwRmSz.Suppress = True    ' manali 3.03
'    'Rep.FldRrRngSz.Suppress = True
 '   Rep.FldqStkRt.Suppress = True
 '   Rep.FldFrRmLoc.Suppress = True
 '   Rep.FldToRmLoc.Suppress = True
 ' ElseIf UCase(adc("UoVchOpt")) = UCase("Vch") Then
    '*** (Bef 15/05/07)
    ''****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] - Company Code added in the Formula Below
    'GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdCoCd})+'/'+({rdo.OdTc})+'/'+" + _
    '"Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))"
    '*** (Bef 15/05/07)
    
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
 '   GRP_REP.SetFormula Rep, "wVchNo", "Trim ({rdo.OdTc})+'/'+" + _
    "Trim ({rdo.OdYy})+'/'+Trim ({rdo.OdChr})+'/'+Trim (ToText ({rdo.OdNo}, '######0'))"
    '*** (Jen 2.14 Next Patch 1) (15/05/07)
  
  '  Rep.hqLoc.Suppress = True
   ' Rep.hRmCtg.Suppress = True
  '  Rep.hRmSCtg.Suppress = True
  '  Rep.hRmCd.Suppress = True
  '  Rep.hqRmSz.Suppress = True
  '  'Rep.hRrRngSz.Suppress = True
  '  Rep.hqStkRt.Suppress = True
  '  Rep.hFrRmLoc.Suppress = True
  '  Rep.hToRmLoc.Suppress = True
  '  Rep.hBagNo.Suppress = True
  '
  '  Rep.FldqLoc.Suppress = True
  '  Rep.FldRmCtg.Suppress = True
  '  Rep.FldRmSCtg.Suppress = True
  '  Rep.FldRmCd.Suppress = True
  '  Rep.FldqRmSz.Suppress = True
  '  Rep.FldwRmSz.Suppress = True    ' manali 3.03
  '  'Rep.FldRrRngSz.Suppress = True
  '  Rep.FldqStkRt.Suppress = True
  '  Rep.FldFrRmLoc.Suppress = True
  '  Rep.FldToRmLoc.Suppress = True
  '  Rep.FldBagNo.Suppress = True
  '  ' Zubin 212
  '  Rep.hLotNo.Suppress = True
  '  Rep.FldLotNo.Suppress = True
  '  ' Zubin 212
 ' End If
   
  '2.04-3 urmi RmWt/Qty Closing Val not Proper so change the Selection Formula
  ' also thr Det Section Suppress fromula has been changed
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  If adc("UoPMCdFr") <> "" Then
'  If adc("UoYN1") = "Y" Then
' ###########################################  Manoj  ###########################################
    '****** (Jen 2.12 Patch 1) put the additional check " Or {rdo.qVal} <> 0 " in the RecordSelectionFormula ******'
      Rep.RecordSelectionFormula = "{rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 Or {rdo.qVal} <> 0 " + _
                                     "Or Round({rdo.qDrWt}, 3) <> 0 Or Round({rdo.qCrWt}, 3) <> 0 " + _
                                     "Or Round({rdo.qOpnWt}, 3) <> 0 Or Round({rdo.qPureWt}, 3)<> 0 " + _
                                     "Or ({rdo.qDrQty} - {rdo.qCrQty} + {rdo.qOpnQty}) <> 0 "
      '*** (Jen 2.13)
      
'    End If
'  End If
 '********************************************************************************
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    Rep.HLotNo.Suppress = True: Rep.FldLotNo.Suppress = True
  End If
  ' Zubin 212
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 7)
  
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
'  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                     "Else (If {rdo.qPtrYN}= 'Y' then ToText({rdo.qRmSz}, 4) Else ToText({rdo.qRmSz}, 3))"
  
  
  ' ****** Manali 3.03 - 09/07/08 - RmSzDesc
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub

errhld:
    moCn.Execute "Delete From #TmpLRmStkLed"
    DispMsg Err.Description, etError
    adc.RepRecSource = "Select * From #TmpLRmStkLed"
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 6    ' Tag = "ShowVal"
    If ms_PwdFlag = False Then
'    If Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
      .Value = Unchecked: DispMsg "Enter SEO Password to View Value ", etError: Exit Sub
    End If
    If .Value = Checked Then adc("UOYN6") = "Y" Else adc("UOYN6") = "N"
 Case Is = 2     'Tag = "RevLogic"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
'   Case Is = 4     'Tag = "ShowDesc"
'    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  Case Is = 0     'Tag = "Show Bal"
   
  'If Not (moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(adc("UODTTO")) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(adc("UODTTO")) + "') as Varchar(4)))),3)='" + CStr(adc("UODTTO")) + "'") Or CInt(Format(CStr(adc("UODTFr")), "dd") = 1)) Or
'  If Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'     .Value = Unchecked:  Exit Sub
'  End If
  
  If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  End Select
  
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

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
  gb_CoCdFor = False
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Else
    Set adc.FirNKeyCtl = adc("UoLocTypFr")
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  'Set adc.FirNKeyCtl = adc("UoLocTypFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = True
  gb_RmZ = False
  'gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  'gs_OmTcTyp = "SO"
  gs_CmCtg = "C"
  gs_Tbl = "Param"
  gs_PTyp = "WIPVAL"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    LblLotNo.Visible = False: adc("UoLotNoFr").Visible = False
    adc("UoLotNoTo").Visible = False: adc("UoLotNoSel").Visible = False
    adc("UoLotNoFr") = "": adc("UoLotNoTo") = "": adc("UoLotNoSel") = ""
  End If
  ' Zubin 212
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: adc("UOPRNTCDFR").Visible = True: adc("UOPRNTCDTO").Visible = True: adc("UOPRNTCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: adc("UOPRNTCDFR").Visible = False: adc("UOPRNTCDTO").Visible = False: adc("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call SetGroupSort
  '*** Jay 2.14 *** (SEO Change)
  ms_PwdFlag = moCn.RecSeek("Select PValue1 from Param where PTyp= 'SEOPWD' and PmCd= 'SEOPWD' and PValue1= '' ")
  '*** Jay 2.14 *** (SEO Change)

  '4.1.0.0
  'ChkBoxArr(2).Visible = False

End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim ws_NewVal As String
  gb_RmCtgFor = True
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr"), UCase("UoPmCdTo")                                                '4.1.4
    If UCase(IdName) = UCase("UoPmCdFr") And pv_NewValue <> "STKVAL" Then adc("UoPMCdTo") = ""  '4.1.4
    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
  
  Case Is = UCase("UoLocTypFr")
    ws_NewVal = pv_NewValue
    gs_LocTyp = IIF(ws_NewVal = "", "", "'" + ws_NewVal + "'")
  
  End Select
  
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
gb_RmCtgFor = True
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
'*** Jay 2.14 *** (SEO Change)
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'    If TxtPwd.text = "" Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
'    If adc("UoYn1") = "N" Then Cancel = True: ErrMsg = "Cannot edit this value as Show Val Option is 'N' ": Exit Sub
' ###########################################  Manoj  ###########################################
    If ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
'*** Jay 2.14 *** (SEO Change)

    gs_Tbl = "Param"
    gs_PTyp = "WIPVAL"
    gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  
  '*** (Jen 2.13)
  Case Is = UCase("UOYN6")
    '*** (Bef 2.14 Next)
    'If adc("UoPMCdFr") <> "" Or adc("UoYN3") = "Y" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    '*** (Bef 2.14 Next)
    '*** (Jen 2.14 Next)
    If ms_PwdFlag = False Then Cancel = True: ErrMsg = "Enter SEO Password to View Value ": Exit Sub
'    If adc("UoPMCdFr") <> "" Then Cancel = True: ErrMsg = "Cannot Edit This Value": Exit Sub
    '*** (Jen 2.14 Next)
  '*** (Jen 2.13)
  '*** Jay 3.01 **** [LotNo Err]
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocFr")
  Case Is = UCase("UoLotNoTo")
    gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocTo")
  ' ******* Manali 350Nxt
  Case Is = UCase("UoLotNoSel")
    'gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocFr")
  ' ******* Manali 350Nxt
  '*** Jay 3.01 **** [LotNo Err]
  
  Case Is = UCase("UoPMCdTo")     '4.1.4
      If adc.Mode = xNorm Then _
        If UCase(adc("UoPMCdFr")) <> UCase("STKVAL") Then Cancel = True: ErrMsg = "Can Enter Only if Value Type is 'STKVAL'": Exit Sub
  
  End Select
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** (Jen 2.13)
'  Select Case UCase(IdName)
'  Case Is = UCase("UoYN")
'    If adc("UoPMCdFr") = "" Then
      'adc("UoYN") = IIF(adc("UoYN3") = "N", "Y", "N")    '*** (Bef 2.14 Next)
'      adc("UoYN") = "Y"                                  '*** (Jen 2.14 Next)
'      OptLastCol(0).Value = True
'    Else
'      adc("UoYN") = "N"
'      OptLastCol(1).Value = True
'    End If
    
'  Case Is = UCase("UoYN5")      '4.1.0.0
'    If Not (moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(adc("UODTTO")) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(adc("UODTTO")) + "') as Varchar(4)))),3)='" + CStr(adc("UODTTO")) + "'") Or CInt(Format(CStr(adc("UODTFr")), "dd") = 1)) Then
'        adc("UoYn5") = "N"
'        ChkBoxArr(0).Value = Unchecked
'    End If
    
'  End Select
  '*** (Jen 2.13)
End Sub


' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptLastCol_Click(Index As Integer)
Select Case Index
Case Is = 0
'  If adc("UoPMCdFr") <> "" Then
'    adc("UoYN") = "N": OptLastCol(1).Value = True: DispMsg "Value Type Is Selected", etError:
'  Else
'    adc("UoYN") = "Y"
'  End If
'Case Is = 1
'  adc("UoYN") = "N"
'  If adc("UoPMCdFr") = "" Then _
'    OptLastCol(2).Value = True: DispMsg "Value Type Is Blank", etError: Exit Sub
'Case Is = 2
'  adc("UoYN") = "N"
'  If adc("UoPMCdFr") <> "" Then
'    OptLastCol(1).Value = True: DispMsg "Value Type Is Selected", etError: Exit Sub
'  Else
'    OptLastCol(2).Value = True
'  End If
End Select
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added
Private Sub OptLastCol_GotFocus(Index As Integer)
'  DispMsg FraOptLastCol.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptQW_Click(Index As Integer)
Select Case Index
Case Is = 0
   adc("UoYN3") = "Y"
Case Is = 1
   adc("UoYN3") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptQW_GotFocus(Index As Integer)
  DispMsg FraOptQW.ToolTipText, etInfo
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  Call HideAllFras
' ###########################################  Manoj  ###########################################
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)

  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

 ' Call SetGroupSort     '4.1.4

  'adc("UoCoCdFr") = gs_CoCd
  adc("UoOpnYyMm") = 0
  gs_LocTyp = "'" + adc("UoLocTypFr") + "'"
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'  txtPwd.text = ""
'  adc("UoPMCdFr") = ""
  adc("UoYn5") = "N"    '4.1.0.0
  
'  adc("UOYN1") = "N"
'  adc("UOYN4") = "N"
' ###########################################  Manoj  ###########################################

  adc("UOYN2") = "Y"  '*** (Jen 3.01)
  adc("UOYN6") = "N"
  
  ' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then OptQW(0).Value = True
  If adc("UoYN3") = "N" Then OptQW(1).Value = True
  
  If adc("UOYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  'If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  'If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  'If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
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
  
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  
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
' ########################  Manoj #### Ver: 2.0.6 #### Date: 24/01/2004  ########################
'**** Shilpa Emr 206 30/10/03 *****/
'Private Sub CmdPwd_Click()
'  Call DispFra(Pwd)
'End Sub

'Private Sub DispFra(ByVal pv_StkLdgFra As en_StkLdgFra)
'  '*** This is the common procedure which is used to toggle between hiding and displaying
'      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
'      'associated with a command button
'  Select Case pv_StkLdgFra
'  Case Is = Pwd
'    If FraPwd.Visible = True Then
'      FraPwd.Visible = False
'      FraPwd.Enabled = False
'      Call EnaDisaCmds(False, CmdPwd)
'      CmdPwd.SetFocus
'    Else
'      FraPwd.Visible = True
'      FraPwd.Enabled = True
'      CmdPwd.SetFocus
'      FraPwd.ZOrder
'      Call EnaDisaCmds(True, CmdPwd)
'    End If
'  End Select
'End Sub
' ###########################################  Manoj  ###########################################
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As CommandButton)
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
      pv_CmdBut.BackColor = vbWhite
      pv_CmdBut.FontBold = True
      pv_CmdBut.CausesValidation = True
    Else
      pv_CmdBut.BackColor = &H8000000F
      pv_CmdBut.FontBold = False
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub

Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  If Trim(TxtPwd.text) = "" Then
      'adc("UoPMCdFr") = ""
      If moCn.GetFldVal("Select PValue1 from Param where Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") = "" Then _
         ms_PwdFlag = True Else ms_PwdFlag = False
  Else
      If Trim(TxtPwd.text) = moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                     "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                     "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                     "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ") Then
          ms_PwdFlag = True: Exit Sub
      Else
          ms_PwdFlag = False
          Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
      End If
  End If
End Sub
