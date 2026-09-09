VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpStkSumm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Summary"
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
      Left            =   12180
      TabIndex        =   44
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
      TabIndex        =   45
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
      Left            =   6135
      TabIndex        =   54
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   55
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
         TabIndex        =   56
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
      Left            =   7350
      TabIndex        =   53
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   50
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
      TabIndex        =   46
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
      Height          =   9765
      Left            =   0
      TabIndex        =   49
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   48
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpStkSumm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpStkSumm.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   52
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   47
               Top             =   360
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   51
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
               Height          =   585
               Index           =   2
               Left            =   0
               TabIndex        =   79
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Stock Rate"
               Top             =   6030
               Width           =   1740
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1515
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   43
               Top             =   5670
               Width           =   2340
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Custom Pure Wt"
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
               Left            =   3510
               TabIndex        =   42
               Tag             =   "ShowZWt"
               ToolTipText     =   "Check To Show ZPurity Wt"
               Top             =   5355
               Width           =   2505
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Pure Wt "
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
               Left            =   30
               TabIndex        =   41
               Tag             =   "ShowPureWt"
               ToolTipText     =   "Check To Show Pure Weight"
               Top             =   5355
               Width           =   1680
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   3915
               TabIndex        =   9
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1635
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "Base.LocCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6300
               TabIndex        =   10
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1635
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "Base.LocCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   3915
               TabIndex        =   38
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   4785
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   37
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   4785
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1515
               TabIndex        =   8
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1635
               Width           =   1490
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "Base.LocCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3915
               TabIndex        =   22
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   3075
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
               Index           =   15
               Left            =   6300
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   3075
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
               Index           =   4
               Left            =   3915
               TabIndex        =   16
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   2505
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
               Index           =   5
               Left            =   6300
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   2505
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
               Index           =   6
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   1920
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1515
               TabIndex        =   15
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   2505
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
               Index           =   9
               Left            =   1515
               TabIndex        =   21
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   3075
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
               TabIndex        =   34
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   4215
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "TdRmSz <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1515
               TabIndex        =   33
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   4215
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "TdRmSz >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   3915
               TabIndex        =   36
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   4500
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "TdRmStkRt<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1515
               TabIndex        =   35
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   4500
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "TdRmStkRt>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1515
               TabIndex        =   3
               ToolTipText     =   "Enter From Location Type"
               Top             =   1050
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
               CmpStr          =   "Base.LocTyp>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   3915
               TabIndex        =   4
               ToolTipText     =   "Enter To Location Type"
               Top             =   1050
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPTO"
               IdName          =   "UOLOCTYPTO"
               CmpStr          =   "Base.LocTyp<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1995
               TabIndex        =   40
               ToolTipText     =   "Show Pure Weight (Yes / No)"
               Top             =   5355
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
               Index           =   13
               Left            =   3915
               TabIndex        =   31
               ToolTipText     =   "Enter To Raw Material Range Size"
               Top             =   3930
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
               Index           =   17
               Left            =   6300
               TabIndex        =   32
               ToolTipText     =   "Enter Raw Material Range Size Selection"
               Top             =   3930
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
               Index           =   19
               Left            =   1515
               TabIndex        =   30
               ToolTipText     =   "Enter From Raw Material Range Size"
               Top             =   3930
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
               TabIndex        =   19
               ToolTipText     =   "Enter To Parent Code"
               Top             =   2790
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
               Index           =   21
               Left            =   6300
               TabIndex        =   20
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   2790
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
               Index           =   22
               Left            =   1515
               TabIndex        =   18
               ToolTipText     =   "Enter From Parent Code"
               Top             =   2790
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
               Top             =   765
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "Base.LocCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   39
               ToolTipText     =   "Enter Open YearMonth"
               Top             =   5070
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
               Left            =   3915
               TabIndex        =   28
               ToolTipText     =   "Enter RM Base Code To Range"
               Top             =   3645
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UOBASECDTO"
               IdName          =   "UOBASECDTO"
               CmpStr          =   "RmBaseCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   6300
               TabIndex        =   29
               ToolTipText     =   "Enter RM Base Code Selection"
               Top             =   3645
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOBASECDSEL"
               IdName          =   "UOBASECDSEL"
               CmpStr          =   "RmBaseCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1515
               TabIndex        =   27
               ToolTipText     =   "Enter RM Base Code From Range"
               Top             =   3645
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UOBASECDFR"
               IdName          =   "UOBASECDFR"
               CmpStr          =   "RmBaseCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   1515
               TabIndex        =   24
               ToolTipText     =   "Enter From Lot No."
               Top             =   3360
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   3915
               TabIndex        =   25
               ToolTipText     =   "Enter To Lot No."
               Top             =   3360
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6300
               TabIndex        =   26
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3360
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
               Index           =   61
               Left            =   3915
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Stone Class To Range"
               Top             =   2220
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "IsNull(PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   6300
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   2220
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "IsNull(PSCd,'') In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   1515
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Stone Class From Range"
               Top             =   2220
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "IsNull(PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   3915
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   750
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "Base.LocCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6300
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   750
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "Base.LocCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   3915
               TabIndex        =   6
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1350
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDTO"
               IdName          =   "UOPRNTCDTO"
               CmpStr          =   "Base.LocPrntCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   6300
               TabIndex        =   7
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1350
               Width           =   8700
               _ExtentX        =   15346
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "Base.LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1515
               TabIndex        =   5
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1350
               Width           =   1490
               _ExtentX        =   2619
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPRNTCDFR"
               IdName          =   "UOPRNTCDFR"
               CmpStr          =   "Base.LocPrntCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   6240
               TabIndex        =   77
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   5400
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   240
               Index           =   71
               Left            =   2400
               TabIndex        =   80
               ToolTipText     =   "Show Custom Pure Wt? (Y/N)"
               Top             =   6240
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   423
               Enabled         =   0   'False
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
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
               Left            =   0
               TabIndex        =   81
               Top             =   5670
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "OR"
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
               Left            =   2760
               TabIndex        =   78
               Top             =   5355
               Width           =   465
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
               TabIndex        =   76
               Top             =   1350
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
               Left            =   30
               TabIndex        =   75
               Top             =   2235
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
               TabIndex        =   74
               Top             =   3375
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "BaseRmCode"
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
               TabIndex        =   73
               Top             =   3660
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
               TabIndex        =   72
               Top             =   5085
               Width           =   1425
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
               TabIndex        =   71
               Top             =   765
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
               TabIndex        =   70
               Top             =   2805
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
               TabIndex        =   69
               Top             =   3945
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Pure Wt"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   10
               Left            =   60
               TabIndex        =   68
               ToolTipText     =   "Location"
               Top             =   5370
               Visible         =   0   'False
               Width           =   1485
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
               TabIndex        =   67
               Top             =   1050
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
               TabIndex        =   66
               Top             =   4230
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
               TabIndex        =   65
               ToolTipText     =   "Location"
               Top             =   1920
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
               TabIndex        =   64
               Top             =   2520
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
               TabIndex        =   63
               Top             =   3090
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Stock Rate"
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
               TabIndex        =   62
               Top             =   4515
               Width           =   1425
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
               TabIndex        =   61
               Top             =   4800
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               TabIndex        =   60
               Top             =   1635
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
               TabIndex        =   59
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
               TabIndex        =   58
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
               TabIndex        =   57
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpStkSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepStkSumm
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
'  wLocPos1 = "str(Case max(Base.LocTyp) when 'Xp' then 1 when 'F' then 2 when 'Xc' then 3 When 'Xl' then 4 when 'W' then 5 when 'S' then 6 when 'L' then 7 When 'B' then 8 when 'M' then 9 else 0 end)"
'  wLocPos2 = "str(Case Base.LocTyp when 'Xp' then 1 when 'F' then 2 when 'Xc' then 3 When 'Xl' then 4 when 'W' then 5 when 'S' then 6 when 'L' then 7 When 'B' then 8 when 'M' then 9 else 0 end)"
  
  Dim wLocPos1 As String, wLocPos2 As String
  wLocPos1 = "str(Case max(Base.LocTyp) when 'S' then 1 when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end)"
  wLocPos2 = "str(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end)"
  
  GRP_REP.Add "Loc Type", Array(wLocPos1 + "+max(Base.LocTyp)", wLocPos1 + "+max(Base.LocTyp)", wLocPos2 + "+Base.LocTyp", wLocPos1 + "+max(Base.LocTyp)", wLocPos2 + "+Base.LocTyp", wLocPos1 + "+max(Base.LocTyp)", wLocPos2 + "+Base.LocTyp", wLocPos1 + "+max(Base.LocTyp)", wLocPos2 + "+Base.LocTyp"), _
  Array("max(Base.LocTyp)", "max(Base.LocTyp)", "Base.LocTyp", "max(Base.LocTyp)", "Base.LocTyp", "max(Base.LocTyp)", "Base.LocTyp", "max(Base.LocTyp)", "Base.LocTyp"), _
  "qLocTyp", "hqLocTyp", "", "", "", Array("(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = max(Base.LocTyp)) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = max(Base.LocTyp)) ", _
  "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = Base.LocTyp) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = max(Base.LocTyp)) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = Base.LocTyp) ", _
  "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = max(Base.LocTyp)) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = Base.LocTyp) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = max(Base.LocTyp)) ", "(Select PDesc From Param Where PTyp= 'LOCTYP' and PMCd = Base.LocTyp) ")
  
  'GRP_REP.Add "Process", "LocCoCd+LocCd", "'['+LocCoCd+'] '+LocCd", "qLoc", "hqLoc", "Company Code,Parent Code", "", "", "LocDesc"
  
  '(Bef 10/01/07) GRP_REP.Add "Location", Array(wLocPos1 + "+Base.LocCd", wLocPos1 + "+Base.LocCd", wLocPos2 + "+Base.LocCd", wLocPos1 + "+Base.LocCd", wLocPos2 + "+Base.LocCd", wLocPos1 + "+Base.LocCd", wLocPos2 + "+Base.LocCd"), "'['+Base.LocCoCd'] '+Base.LocCd", "qLoc", "hqLoc", "Company Code", "", "", Array("max(Base.LocDesc)", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc")
  'GRP_REP.Add "Location", Array(wLocPos1 + "+Base.LocCd", wLocPos1 + "Base.LocCoCd+Base.LocCd", wLocPos2 + "'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "Base.LocCoCd+Base.LocCd", wLocPos2 + "'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "Base.LocCoCd+Base.LocCd", wLocPos2 + "'['+Base.LocCoCd+']'+Base.LocCd"), "'['+ Base.LocCoCd+ '] '+Base.LocCd", "qLoc", "hqLoc", "Company Code", "", "", Array("max(Base.LocDesc)", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc")
  
  'GRP_REP.Add "Location", Array(wLocPos1 + "+Base.LocCoCd+Base.LocCd", wLocPos1 + "+Base.LocCoCd+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+Base.LocCoCd+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+Base.LocCoCd+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd"), "+'['+Base.LocCoCd+']'+Base.LocCd", "qLoc", "hqLoc", "Company Code", "", "", Array("max(Base.LocDesc)", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc")
  
  ' ****** Manali - 21/02/09 - Error rectified
  GRP_REP.Add "Location", Array(wLocPos1 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos1 + "+'['+Base.LocCoCd+']'+Base.LocCd", wLocPos2 + "+'['+Base.LocCoCd+']'+Base.LocCd"), "+'['+Base.LocCoCd+']'+Base.LocCd", "qLoc", "hqLoc", "Company Code", "", "", Array("max(Base.LocDesc)", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc", "max(Base.LocDesc)", "Base.LocDesc")
  
  
  'GRP_REP.Add "Rm Ctg", "RmCtg", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(RmSCtg)", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg)"
  GRP_REP.Add "Rm Code", "Rtrim(RmSCtg)+Rtrim(RmCd)", "'('+ RmSCtg + ') '+ RmCd", "RmCd", "hRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", Array("Max(SubString(RmDesc, 1, 30))", "Max(SubString(RmDesc, 1, 30))", "SubString(RmDesc, 1, 30)", "Max(SubString(RmDesc, 1, 30))", "SubString(RmDesc, 1, 30)", "Max(SubString(RmDesc, 1, 30))", "SubString(RmDesc, 1, 30)", "Max(SubString(RmDesc, 1, 30))", "SubString(RmDesc, 1, 30)")
  
  'GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  '"'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)"
  '"'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)"
  '"Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)"
  '"Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)"
  '"(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)"
  '"(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))"
 
  GRP_REP.Add "Parent Cd", Array("Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", _
  "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", _
  "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", _
  "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)"), _
  Array("'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", _
  "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", _
  "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when max(RmParentCd)= '' then RmCd else max(RmParentCd) end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)"), _
  "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", _
  Array("(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)", _
  "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd))", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)")
  
  '****************Geeta**************Emr206*********12/04/04
  ' GRP_REP.Add "Rm Sv/Len", Array("str(SRmSz,6,3)", "str(TdRmSz,6,3)", "str(TdRmSz,6,3)", "str(FrRmSz,6,3)", "str(FrRmSz,6,3)", "str(IrRmSz,6,3)", "str(IrRmSz,6,3)"), "", "qRmSz", "hqRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
'  GRP_REP.Add "Rm Sv/Len", _
'  Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end", _
'  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""
  ' ******* Manali 3.03 - RmSzDesc added

'  GRP_REP.Add "Rm Sv/Len", _
  Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(SRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(TdRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(RisRmSzDesc, str(TdRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(IrRmSz,6,3)) end", _
  "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(RisRmSzDesc, str(IrRmSz,6,3)) end", _
  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
  "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end "), "", "wRmSz", "hqRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""
  '*****************
  
  GRP_REP.Add "Rm Sv/Sz", _
  Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end "), _
        "", "wRmSz", "hRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  
'  GRP_REP.Add "Rm Sv/Sz", _
  Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end "), _
        "", "wRmSz", "hRmSz", "Rm Ptr,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
  
  
  GRP_REP.Add "Rm StkRt", Array("str(SRmStkRt,9,2)", "str(TdRmStkRt,9,2)", "str(TdRmStkRt,9,2)", "Str(FrRmStkRt,9,2)", "Str(FrRmStkRt,9,2)", "str(IrStkRt,9,2)", "str(IrStkRt,9,2)", "Str(FrRmStkRt,9,2)", "Str(FrRmStkRt,9,2)"), "", "qRmStkRt", "hqRmStkRt", "Rm Sv/Len,Rm Ptr,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""
  If (moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and " + _
                                                  "HCd='" + ctSelfCmCd + "'") = "Y" And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then _
     ' GRP_REP.Add "Rm Ptr", Array("str(SRmSz,6,3)", "str(TdRmSz,6,3)", "str(TdRmSz,6,3)", "str(FrRmSz,6,3)", "str(FrRmSz,6,3)", "str(IrRmSz,6,3)", "str(IrRmSz,6,3)"), "", "qRmSz", "hqRmSz", "Rm Sv/Len,Rm StkRt,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""
     '****************Geeta**************Emr206*********12/04/04
'     GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end", _
'      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else str(IrRmSz,6,3) end"), "", "qRmSz", "hqRmSz", "Rm Sv/Len,Rm StkRt,(RmCd) Sv/Len,(RmCd) StkRt", "", "", ""

     GRP_REP.Add "Rm Ptr", Array("case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(SRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(TdRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(RisRmSzDesc, str(TdRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(IrRmSz,6,3)) end", _
      "case when " + wsPtrYN + "='Y' then   str(IrRmSz,7,4) else IsNull(RisRmSzDesc, str(IrRmSz,6,3)) end", _
      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
      "case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end "), "", "wRmSz", "hRmSz", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
      
'      GRP_REP.Add "Rm Ptr", _
      Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end "), _
        "", "wRmSz", "hqRmSz", "Rm Sv/Sz,Rm StkRt,(RmCd) Sv/Sz,(RmCd) StkRt", "", "", ""
      '*****************
  End If
  GRP_REP.Add "Range Size", Array("RmSCtg+max(IsNull(RrRngSz, ''))", "RmSCtg+max(IsNull(RrRngSz, ''))", "RmSCtg+IsNull(RrRngSz, '')", "RmSCtg+max(IsNull(RrRngSz, ''))", "RmSCtg+IsNull(RrRngSz, '')", "RmSCtg+max(IsNull(RrRngSz, ''))", "RmSCtg+IsNull(RrRngSz, '')", "RmSCtg+max(IsNull(RrRngSz, ''))", "RmSCtg+IsNull(RrRngSz, '')"), Array("'('+ RmSCtg+ ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "'('+ RmSCtg+ ') '+max(IsNull(RrRngSz, ''))", _
    "'('+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "'('+ RmSCtg+ ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ ') '+IsNull(RrRngSz, '')", "'('+ RmSCtg+ ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ ') '+IsNull(RrRngSz, '')"), "RmSCtg", "hRmSCtg", "Rm SubCtg", "", "", ""
  
  'GRP_REP.Add "(RmCd) Sv/Len", Array("RmCd+str(SRmSz,6,3)", "RmCd+str(TdRmSz,6,3)", _
  "RmCd+str(TdRmSz,6,3)", "RmCd+str(FrRmSz,6,3)", "RmCd+str(FrRmSz,6,3)", _
  "RmCd+str(IrRmSz,6,3)", "RmCd+str(IrRmSz,6,3)"), Array("'('+RmCd+') '+LTrim(str(SRmSz,6,3))", "'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(FrRmSz,6,3))", "'('+RmCd+') '+LTrim(str(FrRmSz,6,3))", _
  "'('+RmCd+') '+LTrim(str(IrRmSz,6,3))", "'('+RmCd+') '+LTrim(str(IrRmSz,6,3))"), "RmCtg,qRmSz", "hRmCtg,hqRmSz", _
  "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
  
  '****************Geeta**************Emr206*********12/04/04
'  GRP_REP.Add "(RmCd) Sv/Len", _
'     Array("RmCd+ case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else str(SRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else str(TdRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else str(FrRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else str(IrRmSz,6,3) end ", _
'     "RmCd+ case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else str(IrRmSz,6,3) end "), _
'     Array("'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then Ltrim(str(SRmSz,7,4)) else Ltrim(str(SRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(TdRmSz,7,4)) else Ltrim(str(TdRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(TdRmSz,7,4)) else Ltrim(str(TdRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(str(FrRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(str(FrRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(IrRmSz,7,4)) else LTrim(str(IrRmSz,6,3)) end ", _
'     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(IrRmSz,7,4)) else LTrim(str(IrRmSz,6,3)) end "), _
'     "RmCtg,qRmSz", "hRmCtg,hqRmSz", _
'     "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
   ' ******* Manali 3.03 - RmSzDesc added

'   GRP_REP.Add "(RmCd) Sv/Len", _
     Array("RmCd+ case when " + wsPtrYN + "='Y' then str(SRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(SRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(TdRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(TdRmSz,7,4) else IsNull(RisRmSzDesc, str(TdRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(IrRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(IrRmSz,7,4) else IsNull(RisRmSzDesc, str(IrRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3)) end ", _
     "RmCd+ case when " + wsPtrYN + "='Y' then  str(FrRmSz,7,4) else IsNull(RisRmSzDesc, str(FrRmSz,6,3)) end "), _
     Array("'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then Ltrim(str(SRmSz,7,4)) else Ltrim(IsNull(Max(RisRmSzDesc), str(SRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(TdRmSz,7,4)) else Ltrim(IsNull(Max(RisRmSzDesc), str(TdRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(TdRmSz,7,4)) else Ltrim(IsNull(RisRmSzDesc, str(TdRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(IsNull(RisRmSzDesc, str(FrRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(IrRmSz,7,4)) else LTrim(IsNull(Max(RisRmSzDesc), str(IrRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(IrRmSz,7,4)) else LTrim(IsNull(RisRmSzDesc, str(IrRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(IsNull(Max(RisRmSzDesc), str(FrRmSz,6,3))) end ", _
     "'('+RmCd+')'+  case when " + wsPtrYN + "='Y' then  Ltrim(str(FrRmSz,7,4)) else LTrim(IsNull(RisRmSzDesc, str(FrRmSz,6,3))) end "), _
     "RmCtg,wRmSz", "hRmCtg,hqRmSz", _
     "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
      '*****************
  
'     GRP_REP.Add "(RmCd) Sv/Sz", _
      Array("RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end "), _
      Array("'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(Max(RisRmSzDesc), " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end ", _
        "'('+RmCd+')'+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else IsNull(RisRmSzDesc, " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + ") end "), _
        "RmCtg,wRmSz", "hRmCtg,hRmSz", _
        "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
     
      GRP_REP.Add "(RmCd) Sv/Sz", _
      Array("RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "RmCd+ case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end "), _
      Array("'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDispFldNm("SRmSz", "SRmSz2", "SRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("TdRmSz", "TdRmSz2", "TdRmSz3") + " else RisRmSzDesc end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("IrRmSz", "IrRmSz2", "IrRmSz3") + " else RisRmSzDesc end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when Max(RisRmSzDesc) is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else Max(RisRmSzDesc) end) end ", _
        "'('+RmCd+') '+ case when " + wsPtrYN + "='Y' then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else (case when RisRmSzDesc is null then " + GetRmSzDispFldNm("FrRmSz", "FrRmSz2", "FrRmSz3") + " else RisRmSzDesc end) end "), _
        "RmCd,wRmSz", "hRmCd,hRmSz", _
        "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
     
     
    '****** Sachin 2.13.0 - [07-07-2006] ******
    GRP_REP.Add "Stone Class", _
    Array("RmCtg+Max(IsNull(PSCd, ''))", "RmCtg+Max(IsNull(PSCd, ''))", "RmCtg+IsNull(PSCd, '')", _
          "RmCtg+Max(IsNull(PSCd, ''))", "RmCtg+IsNull(PSCd, '')", "RmCtg+Max(IsNull(PSCd, ''))", _
          "RmCtg+IsNull(PSCd, '')", "RmCtg+Max(IsNull(PSCd, ''))", "RmCtg+IsNull(PSCd, '')"), _
    Array("'('+ RTrim(RmCtg)+') '+Max(IsNull(PSCd, ''))", "'('+ RTrim(RmCtg)+') '+Max(IsNull(PSCd, ''))", _
          "'('+ RTrim(RmCtg)+') '+IsNull(PSCd, '')", "'('+ RTrim(RmCtg)+') '+Max(IsNull(PSCd, ''))", _
          "'('+ RTrim(RmCtg)+') '+IsNull(PSCd, '')", "'('+ RTrim(RmCtg)+') '+Max(IsNull(PSCd, ''))", _
          "'('+ RTrim(RmCtg)+') '+IsNull(PSCd, '')", "'('+ RTrim(RmCtg)+') '+Max(IsNull(PSCd, ''))", _
          "'('+ RTrim(RmCtg)+') '+IsNull(PSCd, '')"), "RmCtg", "", "Rm Ctg", "", "", _
    Array("Max(IsNull(PDesc,''))", "Max(IsNull(PDesc,''))", "IsNull(PDesc,'')", "Max(IsNull(PDesc,''))", _
          "IsNull(PDesc,'')", "Max(IsNull(PDesc,''))", "IsNull(PDesc,'')", "Max(IsNull(PDesc,''))", "IsNull(PDesc,'')")
    '****** Sachin 2.13.0 - [07-07-2006] ******
     
  GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(SRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", _
  "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(IrStkRt,9,2)", "RmCd+str(IrStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)"), _
  Array("'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(IrStkRt,9,2))", "'('+RmCd+') '+LTrim(str(IrStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))"), _
  "RmCtg,qStkRt", "hRmCtg,hqStkRt", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Sv/Len,Rm Sv/Len,Rm Ptr,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
'  GRP_REP.Add "(RmCd) Sv/Len", Array("RmCd+str(SRmSz,6,3)", "RmCd+str(TdRmSz,6,3)", "RmCd+str(TdRmSz,6,3)", "RmCd+str(FrRmSz,6,3)", "RmCd+str(FrRmSz,6,3)"), Array("'('+RmCd+') '+LTrim(str(SRmSz,6,3))", "'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(TdRmSz,6,3))", "'('+RmCd+') '+LTrim(str(FrRmSz,6,3))", "'('+RmCd+') '+LTrim(str(FrRmSz,6,3))"), "RmCtg,RmSCtg,RmCd,qRmSz", "hRmCtg,hRmSCtg,hRmCd,hqRmSz", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) StkRt,Rm Sv/Len,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "
'  GRP_REP.Add "(RmCd) StkRt", Array("RmCd+str(SRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(TdRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)", "RmCd+str(FrRmStkRt,9,2)"), Array("'('+RmCd+') '+LTrim(str(SRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(TdRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))", "'('+RmCd+') '+LTrim(str(FrRmStkRt,9,2))"), "RmCtg,RmSCtg,RmCd,qStkRt", "hRmCtg,hRmSCtg,hRmCd,hqStkRt", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Sv/Len,Rm Sv/Len,Rm StkRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')' "

  ' Zubin 212
  If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y'") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
  'If moCn.GetFldVal("select Max(HLotNoYn) from Head where HCd= '" + ctSelfCmCd + "' " + IIF(adc("UoCoCdFr") <> "", " and HCoCd>= '" + adc("UoCoCdFr") + "' ", " ") + IIF(adc("UoCoCdTo") <> "", " and HCoCd<= '" + adc("UoCoCdTo") + "' ", " ") + IIF(adc("UoCoCdSel") <> "", " and HCoCd In (" + adc("UoCoCdSel") + ") ", " ")) = "Y" Then
    GRP_REP.Add "RM Lot No.", Array("SLotNo", "TdLotNo", "TdLotNo", "FrLotNo", "FrLotNo", "IrLotNo", "IrLotNo", "FrLotNo", "FrLotNo"), "", "qLotNo", "hLotNo", "", "", "", Array("SLotNo", "TdLotNo", "TdLotNo", "FrLotNo", "FrLotNo", "IrLotNo", "IrLotNo", "FrLotNo", "FrLotNo")
  End If
  ' Zubin 212
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "Base.LocCoCd", "Base.LocCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=Base.LocCoCd) "
    GRP_REP.Add "Parent Location", "Base.LocPrntCd", "Base.LocPrntCd", "", "", "", "", "", "(Select LocDesc from Loc Ln where Ln.LocCoCd= '" + ctSelfCoCd + "' And Ln.LocCd=Base.LocPrntCd) "
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] - ******

End Sub

Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
    If UCase(adc("UoYN")) = UCase("Y") And UCase(adc("UoYN1")) = UCase("Y") Then
        pr_Cancel = True
        pr_ErrMsg = "Select Either Pure wt OR Custom Pure Wt, Not Both."
    End If
End Sub

Private Sub adc_setreprecsource()
' **** Zubin 211 **** '
' **** All Queries are inserted into temporary table #TmpStkSummLed and the Unions are removed **** '
' **** Option (MaxDOP 1) is used in the queries **** '
On Error GoTo ErrHndlr
' **** Zubin 211 **** '

  '*** Report Sql ***
' Zubin 212 changed query
'*** (Jen 2.14)
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RmCtg, PSCd As qStnCls, RmSCtg,
'  TdRmStkRt as qRmStkRt, PValue3 as qPtrYN, TdCoCd, qLocPrntCd, TdFrRmLoc as qLoc, 0.0 as qOpn,
'  0.0 as qPr, 0.0 as qFg, 0.0 as qCnv, 0.0 as qFLs, 0.0 as qWip, 0.0 as qS,
'  0.0 as qPrdLsDr, 0.0 as qPrdLsCr, 0.0 as qBrk, 0.0 as qMsg, 0.0 as qSal,
'  0.0 as qCls, LocTyp as qLocTyp, 0 as qBaseLocTypPos
'  From Txnd, RmMst, Txn, Loc, Param where 1=2
'*** (Jen 2.14)
  
  '*** (Bef 2.14)
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RmCtg, PSCd As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt,
'  PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd as qCoCd, TdTc as qTc, TdYy as qYy,
'  TdChr as qChr, TdNo as qNo, TdSr as qSr, TdSrNo as qSrNo, TdBYy as qBYy,
'  TdBChr as qBChr, TdBNo as qBNo, TDt as qDt,
'  TdFrRmLoc as qLoc, 0.0 as qOpn, 0.0 as qPr, 0.0 as qFg, 0.0 as qCnv, 0.0 as qFLs,
'  0.0 as qWip, 0.0 as qS, 0.0 as qPrdLsDr, 0.0 as qPrdLsCr, 0.0 as qBrk, 0.0 as qMsg, 0.0 as qSal,
'  0 as qMul, 0.0 as qCls, LocTyp as qLocTyp, 0 as qBaseLocTypPos
'  From Txnd, RmMst, Txn, Loc, Param where 1=2
  '*** (Bef 2.14)
  
  '*** Report Sql ***

  Dim grpflds0 As String, grpflds1 As String, grpflds2 As String, grpflds3 As String
  Dim grpflds4 As String, grpflds5 As String, grpflds6 As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wPurityWt As String
  Dim wDtCndOpnYyMm As String, wsqlOpnYyMm As String, wFrDt As String, wDtCndOpnTd As String
  Dim wDtCndDet As String, wSqlOpnTdFr As String, wSqlOpnTdTo As String, wSqlDetFr1 As String
  Dim wSqlDetFr2 As String, wSqlDetFr As String, wSqlDetTo1 As String, wSqlDetTo2 As String
  Dim wSqlDetTo As String, wFrDt1 As String, wDtCndOpnFg As String, wSqlOpnFgFr As String
  Dim wSqlOpnFgTo As String, wDtCndDetFg As String, wSqlDetFgFr As String, wSqlDetFgFr1 As String
  Dim wSqlDetFgFr2 As String, wSqlDetFgTo1 As String, wSqlDetFgTo2 As String, wSqlDetFgTo As String
  Dim wPureWtOrZWt As String, wXPureWtOrZwt As String
  Dim wPureWtOrZWt1 As String, wXPureWtOrZwt1 As String
  
  Dim wDtCndOpnIn As String, wSqlOpnInFr As String, wSqlOpnInTo As String
  Dim wDtCndDetIn As String, wSqlDetInFr As String, wSqlDetInFr1 As String
  Dim wSqlDetInFr2 As String, wSqlDetInTo1 As String, wSqlDetInTo2 As String, wSqlDetInTo As String
  
  ' ***** Manali 3.7.1
  Dim wDtCndOpnXd As String, wSqlOpnXdFr As String, wSqlOpnXdTo As String
  Dim wDtCndDetXd As String, wSqlDetXdFr As String, wSqlDetXdFr1 As String
  Dim wSqlDetXdFr2 As String, wSqlDetXdTo1 As String, wSqlDetXdTo2 As String, wSqlDetXdTo As String
  Dim grpflds7 As String, grpflds8 As String
  ' ***** Manali 3.7.1
  
  Dim wFgStr As String, wInStr As String
  Dim wPurityWt1 As String, wOpnYyStr As String
  
  ' ******** Manali 3.03 - 09/07/08 - RmSzDesc
  Dim wRmSzDescOpn As String, wRmSzDescDet As String
  ' ******** Manali 3.03 - 09/07/08 - RmSzDesc
  
' **** Zubin 211 **** '
  Dim wSetTmpTab As String
' **** Zubin 211 **** '
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0 - New Fields in Where Clause
  
  '****** Sachin 3.02.0 - Check on YYYYMM Entered by user ******
  Dim ws_PurgeDt As String
  'sv.9 getting max(PurgeDate) from vParam for given company code
  ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                          " where vPCoCd >='" + adc("UoCoCdFr") + "' and " + _
                          " vPCoCd <='" + IIF(adc("UoCoCdTo") <> "", adc("UoCoCdTo"), "ZZZ") + "'" + _
                          IIF(adc("UoCoCdSel") <> "", " and vPCoCd In(" + adc("UoCoCdSel") + ")", "") + _
                          " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And ((CDate(adc("UOOPNYYMM")) <= Format(CDate(ws_PurgeDt), "YYYYMM") And adc("UOOPNYYMM") <> 0) Or CDate(adc("UODTFR")) <= CDate(ws_PurgeDt) Or CDate(adc("UODTTO")) <= CDate(ws_PurgeDt)) Then
      DispMsg "From/To Date Or OpenYYYYMM Cannot Be Less Than Purge Date [" + CStr(CDate(ws_PurgeDt)) + "]", etError: Exit Sub
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
  grpflds4 = GRP_REP.GrpFldLst(4)
  grpflds5 = GRP_REP.GrpFldLst(5)
  grpflds6 = GRP_REP.GrpFldLst(6)
  grpflds7 = GRP_REP.GrpFldLst(7)
  grpflds8 = GRP_REP.GrpFldLst(8)
  
  If UCase(adc("UoYN1")) = UCase("Y") Then
    Rep.TxtPureWtOrZWt.SetText ("(Custom Pure Wts)")
    wPureWtOrZWt = "RmPurityZ"
    wXPureWtOrZwt = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=Rm.RmBaseCd " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=Rm.RmPrtKey)", ")")
    wPureWtOrZWt1 = "max(RmPurityZ)"
    wXPureWtOrZwt1 = "(Select XZ.RmPurityZ from RmMst XZ where XZ.RmCd=max(Rm.RmBaseCd) " + IIF(gs_Partition = ctCurrPrtn, " And XZ.RmPrtKey=max(Rm.RmPrtKey))", ")")
  ElseIf UCase(adc("UoYN")) = UCase("Y") Then
    Rep.TxtPureWtOrZWt.SetText ("(Pure Wts)")
    wPureWtOrZWt = "RmPurityWt"
    wXPureWtOrZwt = "RmBasePurityWt"
    wPureWtOrZWt1 = "max(RmPurityWt)"
    wXPureWtOrZwt1 = "max(RmBasePurityWt)"
  Else
    Rep.TxtPureWtOrZWt.SetText ("")
    wPureWtOrZWt = " 1 "
    wXPureWtOrZwt = " 1 "
    wPureWtOrZWt1 = " 1 "
    wXPureWtOrZwt1 = " 1 "
  End If

  If adc("UoYN") = "Y" Or adc("UoYN1") = "Y" Then
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    wPurityWt = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                              "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                              IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + "RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '')), 1) " + _
                "else 1 end) "
  
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    wPurityWt1 = " /(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                              "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                              IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) " + _
                "when Rm.RmCtg= 'X' then " + _
                          "IsNull((Select " + wPureWtOrZWt + " from Param join RmMst on " + _
                          IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + "RmCd= PValue where PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmCtg From RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= max(Rm.RmBaseCd)), '')), 1) " + _
                "else 1 end) "
  Else
    wPurityWt = " "
    wPurityWt1 = " "
  End If
    
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc
  wRmSzDescOpn = "(case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(Max(RisRmCtg), '') and PSCd= RmSCtg)= 'Y' then IsNull(max(RisRmSzDesc), '') Else '' End) "
  wRmSzDescDet = "(case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(RisRmCtg, '') and PSCd= RmSCtg)= 'Y' then IsNull(RisRmSzDesc, '') Else '' End) "
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc
    
  adc("UoLn1Fr").CmpStr = "SRmSz>= "
  adc("UoLn1To").CmpStr = "SRmSz<= "
  adc("UoCstRtFr").CmpStr = "SRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "SRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "SLotNo>= "
  adc("UoLotNoTo").CmpStr = "SLotNo<= "
  adc("UoLotNoSel").CmpStr = "SLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wDtCndOpnYyMm = " and SYyMm < " + Format(adc("UoDtFr"), "yyyymm")
   If adc("UoOpnYyMm") <> 0 Then
    wOpnYyStr = " and SyyMm >= " + CStr(adc("UoOpnYyMM")) + ""
  Else
    wOpnYyStr = ""
  End If
  
  ' **** Zubin **** '
  ' **** Setting the opening tmp table **** '
  ' Zubin 212
  '*** (Bef 2.14)
  'wSetTmpTab = "Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3, " + _
               "space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3, " + _
               "space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc, " + _
               "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, RmCtg, PSCd As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
               "PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd as qCoCd, TdTc as qTc, TdYy as qYy, " + _
               "TdChr as qChr, TdNo as qNo, TdSr as qSr, TdSrNo as qSrNo, TdBYy as qBYy, " + _
               "TdBChr as qBChr, TdBNo as qBNo, TDt as qDt, TdFrRmLoc as qLoc, TdRmWt as qOpn, " + _
               "TdRmWt as qPr, TdRmWt as qFg, TdRmWt as qCnv, TdRmWt as qFLs, TdRmWt as qWip, TdRmWt as qS, " + _
               "TdRmWt as qPrdLsDr, TdRmWt as qPrdLsCr, TdRmWt as qBrk, TdRmWt as qMsg, TdRmWt as qSal, " + _
               "0 as qMul, TdRmWt as qCls, LocTyp as qLocTyp, 0 as qBaseLocTypPos Into #TmpStkSummLed " + _
               "From Txnd, RmMst, Txn, Loc, Param where 1=2 Option (MaxDOP 1) "
  '*** (Bef 2.14)
    
  '****** (Jen 2.14) removed fields TdTc as qTc, TdYy as qYy, " + _
  '                  "TdChr as qChr, TdNo as qNo, TdSr as qSr, TdSrNo as qSrNo, TdBYy as qBYy, " + _
  '                  "TdBChr as qBChr, TdBNo as qBNo, TDt as qDt,
               
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  wSetTmpTab = "Select space(60) as Grp1, space(60) as Grp2, space(60) as Grp3, " + _
               "space(60) as DGrp1, space(60) as DGrp2, space(60) as DGrp3, " + _
               "space(60) as G1Desc, space(60) as G2Desc, space(60) as G3Desc, " + _
               "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, RisRmSzDesc as qRmSzDesc, RmCtg, PSCd As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
               "PValue3 as qPtrYN, RmPurityWt as qRmPurityWt, TdCoCd as qCoCd, LocPrntCd as qLocPrntCd, TdFrRmLoc as qLoc, TdRmWt as qOpn, " + _
               "TdRmWt as qPr, TdRmWt as qFg, TdRmWt as qCnv, TdRmWt as qFLs, TdRmWt as qWip, TdRmWt as qS, " + _
               "TdRmWt as qPrdLsDr, TdRmWt as qPrdLsCr, TdRmWt as qBrk, TdRmWt as qMsg, TdRmWt as qSal, " + _
               "0 as qMul, TdRmWt as qCls, LocTyp as qLocTyp, 0 as qBaseLocTypPos Into #TmpStkSummLed " + _
               "From Txnd, RmMst, Txn, Loc, Param, RmIdSz where 1=2 Option (MaxDOP 1) "
  moCn.CreateTmpTable "#TmpStkSummLed", wSetTmpTab
  ' **** Zubin **** '
  
  '****** (Jen 2.14) removed fields '' as qTc, '' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wsqlOpnYyMm = "Insert Into #TmpStkSummLed Select " + grpflds0 + ", " + _
                "'O' as qFlag, RMCd, SLotNo as qLotNo, SRmSz as qRmSz, SRmSz2 as qRmSz2, SRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, SRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, sum(SRmDrWt- SRmCrWt) as qOpn, 0 as qPr, 0 as qFg, 0 as qCnv, " + _
                "0 as qFLs, 0 as qWip, 0 as qS, 0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, " + _
                "0 as qMsg, 0 as qSal, 1 as qMul, sum(SRmDrWt- SRmCrWt) as qCls, max(Base.LocTyp) as qLocTyp, " + _
                "(Case max(Base.LocTyp) when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
                "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 " + _
                "when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From SYyMm NoLock Join RmMst Rm On RmCd= SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey  ", "") + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Join Loc Base On LocCoCd= SCoCd and LocCd= SLoc " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= SRmSz and RrToLn>= SRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=SRmSz " + _
                "Where SLoc<> '' " + wDtCndOpnYyMm + wOpnYyStr + wCnd + IIF(gs_Partition = ctCurrPrtn, " and SPrtKey='" + ctCurrPrtn + "'  ", " and SYyMm<>'0' ") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, SLotNo, SRmSz, SRmSz2, SRmSz3, SRmStkRt Option (MaxDOP 1) "
  
  '      IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, " + _

  moCn.Execute wsqlOpnYyMm
  
'                "Where SLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnYyMm + wCnd + _

'6.1
  wFrDt = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "mm") + "/01", "01/" + Format(adc("UoDtFr"), "mm")) + "/" + Format(adc("UoDtFr"), "yy")
  
  '****** Sachin 3.02.0 - New fields in the Where Clause
  wDtCndOpnTd = " and TDt >= '" + wFrDt + "' And TDt < '" + CStr(adc("UoDtFr")) + "' and TdDt >= '" + wFrDt + "' And TdDt < '" + CStr(adc("UoDtFr")) + "' "
  
  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "TdLotNo>= "
  adc("UoLotNoTo").CmpStr = "TdLotNo<= "
  adc("UoLotNoSel").CmpStr = "TdLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")

  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " TdFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
'****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) removed fields '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  '****** Sachin 3.02 - Id fields in join
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnTdFr = "Insert Into #TmpStkSummLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case TdFrRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case TdFrRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From Txnd NoLock Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey  ", "") + _
                "join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey  ", "") + _
                "Join Loc Base On TdCoCd= LocCoCd and TdFrRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                "Where  " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdFrRmLoc<> '' " + wDtCndOpnTd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt Option (MaxDOP 1) "
  
  '     IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "
  moCn.Execute wSqlOpnTdFr
  
'                "Where TdFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnTd + wCnd + _

  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "TdLotNo>= "
  adc("UoLotNoTo").CmpStr = "TdLotNo<= "
  adc("UoLotNoSel").CmpStr = "TdLotNo In "
  wrepcnd = adc.RepCond
  'wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " TdToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) removed fields '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnTdTo = "Insert Into #TmpStkSummLed Select " + grpflds1 + "," + _
                "'O' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case TdToRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case TdToRmDc When 'D' Then TdRmWt When 'C' Then -TdRmWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From Txnd NoLock Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey  ", "") + _
                "join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey  ", "") + _
                "Join Loc Base On TdCoCd= LocCoCd and TdToRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdToRmLoc<> '' " + wDtCndOpnTd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, TdLotNo, TdRmSz, TdRmSz2, TdRmSz3, TdRmStkRt Option (MaxDOP 1) "
             
'                "Where TdToRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnTd + wCnd + _

  moCn.Execute wSqlOpnTdTo
  
  wDtCndDet = " and TDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' and TdDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"

  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "TdLotNo>= "
  adc("UoLotNoTo").CmpStr = "TdLotNo<= "
  adc("UoLotNoSel").CmpStr = "TdLotNo In "
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
'  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, "")
    wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  ' **** Zubin 211 **** '
  ' Zubin 212
  
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " TdFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) TdTc as qTc, TdYy as qYy, TdChr as qChr, " + _
  '                  "TdNo as qNo, TdSr as qSr, TdSrNo as qSrNo, TdBYy as qBYy, TdBChr as qBChr, " + _
  '                  "TdBNo as qBNo, TDt as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetFr1 = "Insert Into #TmpStkSummLed Select " + grpflds2 + "," + _
               "'T' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then TdRmWt else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then TdRmWt else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then TdRmWt else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then TdRmWt else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then TdRmWt else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then TdRmWt else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and TdFrRmDc='D' then TdRmWt else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and TdFrRmDc='C' then TdRmWt else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then TdRmWt else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then TdRmWt else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then TdRmWt else 0 end) as qSal, " + _
               "(Case when TdFrRmDc='D' then 1 when TdFrRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then TdRmWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
               "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetFr2 = "From Txnd NoLock join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=TdPrtKey  ", "") + _
               "Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= TdCoCd and Base.LocCd= TdFrRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= TdCoCd and Loc.LocCd= TdToRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdFrRmLoc<> '' " + _
               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') and Not (TdFrRmLocTyp= Loc.LocTyp and TdFrRmLocTyp= 'W') " + _
               wDtCndDet + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
  
'               "Where TdFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' and Base.LocTyp<>'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W')" + wDtCndDet + wCnd
' IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "
  wSqlDetFr = wSqlDetFr1 + wSqlDetFr2
  moCn.Execute wSqlDetFr
               
  adc("UoLn1Fr").CmpStr = "TdRmSz>= "
  adc("UoLn1To").CmpStr = "TdRmSz<= "
  adc("UoCstRtFr").CmpStr = "TdRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "TdRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "TdLotNo>= "
  adc("UoLotNoTo").CmpStr = "TdLotNo<= "
  adc("UoLotNoSel").CmpStr = "TdLotNo In "
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
'  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, "")
    wCnd = IIF(wrepcnd <> "", " and TdAck='Y' and " + wrepcnd, " and TdAck='Y' ")
  ' **** Zubin 211 **** '
  ' Zubin 212
  
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " TdRmCtg = '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " TdRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " TdRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " TdRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " TdToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " TdToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) TdTc as qTc, TdYy as qYy, TdChr as qChr, " + _
  '             "TdNo as qNo, TdSr as qSr, TdSrNo as qSrNo, TdBYy as qBYy, TdBChr as qBChr, " + _
  '             "TdBNo as qBNo, TDt as qDt,
               
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetTo1 = "Insert Into #TmpStkSummLed Select " + grpflds2 + "," + _
               "'T' as qFlag, RmCd, TdLotNo as qLotNo, TdRmSz as qRmSz, TdRmSz2 as qRmSz2, TdRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, TdRmStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then TdRmWt else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then TdRmWt else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then TdRmWt else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then TdRmWt else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then TdRmWt else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then TdRmWt else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and TdToRmDc='D' then TdRmWt else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and TdToRmDc='C' then TdRmWt else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then TdRmWt else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then TdRmWt else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then TdRmWt else 0 end) as qSal, " + _
               "(Case when TdToRmDc='D' then 1 when TdToRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then TdRmWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
               "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetTo2 = "From Txnd NoLock join Txn on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=TPrtKey  ", "") + _
               "Join RmMst Rm On TdRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= TdCoCd and Base.LocCd= TdToRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= TdCoCd and Loc.LocCd= TdFrRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= TdRmSz and RrToLn>= TdRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=TdRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey='" + ctCurrPrtn + "' and ", "") + " TdToRmLoc<> '' " + _
               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') And Not (TdToRmLocTyp= Loc.LocTyp and TdToRmLocTyp= 'W')  " + _
               wDtCndDet + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
               
  wSqlDetTo = wSqlDetTo1 + wSqlDetTo2
  moCn.Execute wSqlDetTo
      
'               "Where TdToRmLoc<> '' and substring(Base.LocTyp,1,1)<> 'X' and Base.LocTyp<> 'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') " + wDtCndDet + wCnd
  
  
  '*** If Location 'W' is selected, an additional Unions FgRm FrLoc, FgRm ToLoc is required.
  
  wSqlOpnFgFr = "": wSqlOpnFgTo = "": wSqlDetFgFr = "": wSqlDetFgTo = ""
  wFrDt1 = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "mm") + "/01", "01/" + Format(adc("UoDtFr"), "mm")) + "/" + Format(adc("UoDtFr"), "yy") '6.1
  
  '****** Sachin 3.02.0 - New fields in Where Clause
  wDtCndOpnFg = " and FgDt >= '" + wFrDt1 + "' And FgDt < '" + CStr(adc("UoDtFr")) + "' and FrDt >= '" + wFrDt1 + "' And FrDt < '" + CStr(adc("UoDtFr")) + "' "

  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnFgFr = "Insert Into #TmpStkSummLed Select " + grpflds3 + "," + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case FrFrRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case FrFrRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From FgRm NoLock Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey  ", "") + _
                "join Fg on FgCoCd= FrCoCd and FgTc= FrTc and FgYy= FrYy and FgChr= FrChr and FgNo= FrNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FrPrtKey  ", "") + _
                "Join Loc Base On FrCoCd= LocCoCd and FrFrRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and ", "") + " FrFrRmLoc<> '' " + wDtCndOpnFg + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt Option (MaxDOP 1) "

'                "Where FrFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnFg + wCnd
  moCn.Execute wSqlOpnFgFr
    
  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnFgTo = "Insert Into #TmpStkSummLed Select " + grpflds3 + ", " + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case FrToRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case FrToRmDc When 'D' Then FrRmWt When 'C' Then -FrRmWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From FgRm NoLock Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey  ", "") + _
                "join Fg on FgCoCd= FrCoCd and FgTc= FrTc and FgYy= FrYy " + _
                "and FgChr= FrChr and FgNo= FrNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FrPrtKey  ", "") + _
                "Join Loc Base On FrCoCd= LocCoCd and FrToRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and ", "") + " FrToRmLoc<> '' " + wDtCndOpnFg + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt Option (MaxDOP 1) "

'                "Where FrToRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnFg + wCnd +
'                        IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "
  moCn.Execute wSqlOpnFgTo
  
  '****** Sachin 3.02.0 - New fields in Where Clause
  wDtCndDetFg = " and FgDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' and FrDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"

  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' **** Zubin 211 **** '
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) FrTc as qTc, FrYy as qYy, FrChr as qChr, " + _
  '             "FrNo as qNo, FrSr as qSr, FrSrNo as qSrNo, FdBYy as qBYy, FdBChr as qBChr, " + _
  '             "FdBNo as qBNo, FgDt as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetFgFr1 = "Insert Into #TmpStkSummLed Select " + grpflds4 + "," + _
               "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then FrRmWt else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then FrRmWt else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then FrRmWt else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then FrRmWt else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then FrRmWt else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then FrRmWt else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and FrFrRmDc='D' then FrRmWt else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and FrFrRmDc='C' then FrRmWt else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then FrRmWt else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then FrRmWt else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then FrRmWt else 0 end) as qSal, " + _
               "(Case   when FrFrRmDc='D' then 1 when FrFrRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then FrRmWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetFgFr2 = "From FgRm NoLock " + _
               "join Fgd on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy " + _
               "and FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey  ", "") + _
               "join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey  ", "") + _
               "Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= FrCoCd and Base.LocCd= FrFrRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= FrCoCd and Loc.LocCd= FrToRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and  ", "") + " FrFrRmLoc<> '' " + _
               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') and Not (FrFrRmLocTyp= Loc.LocTyp and FrFrRmLocTyp= 'W')" + _
               wDtCndDetFg + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
  
  wSqlDetFgFr = wSqlDetFgFr1 + wSqlDetFgFr2
  moCn.Execute wSqlDetFgFr
  
'               "Where FrFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' and Base.LocTyp<>'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W')" + wDtCndDetFg + wCnd
'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "

  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' **** Zubin 211 **** '
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) FrTc as qTc, FrYy as qYy, FrChr as qChr, " + _
  '                  "FrNo as qNo, FrSr as qSr, FrSrNo as qSrNo, FdBYy as qBYy, FdBChr as qBChr, " + _
  '                  "FdBNo as qBNo, FgDt as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetFgTo1 = "Insert Into #TmpStkSummLed Select " + grpflds4 + "," + _
                 "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                 "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                 "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
                 "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
                 "(Case when Loc.LocTyp = 'Xp' then FrRmWt else 0 end) as qPr, " + _
                 "(Case when Loc.LocTyp = 'F' then FrRmWt else 0 end) as qFg, " + _
                 "(Case when Loc.LocTyp In ('Xc', 'Xd') then FrRmWt else 0 end) as qCnv, " + _
                 "(Case when Loc.LocTyp = 'Xl' then FrRmWt else 0 end) as qFLs, " + _
                 "(Case when Loc.LocTyp = 'W' then FrRmWt else 0 end) as qWip, " + _
                 "(Case when Loc.LocTyp = 'S' then FrRmWt else 0 end) as qS, " + _
                 "(Case when Loc.LocTyp = 'L' and FrToRmDc='D' then FrRmWt else 0 end) as qPrdLsDr, " + _
                 "(Case when Loc.LocTyp = 'L' and FrToRmDc='C' then FrRmWt else 0 end) as qPrdLsCr, " + _
                 "(Case when Loc.LocTyp = 'B' then FrRmWt else 0 end) as qBrk, " + _
                 "(Case when Loc.LocTyp = 'M' then FrRmWt else 0 end) as qMsg, " + _
                 "(Case when Loc.LocTyp = 'Xs' then FrRmWt else 0 end) as qSal, " + _
                 "(Case when FrToRmDc='D' then 1 when FrToRmDc='C' then -1 else 0 end) as qMul, " + _
                 "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
                 "then FrRmWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
                 "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
                 "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 " + _
                 "when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetFgTo2 = "From FgRm NoLock " + _
               "join Fgd on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy " + _
               "and FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey  ", "") + _
               "join Fg on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo " + IIF(gs_Partition = ctCurrPrtn, " and FgPrtKey=FdPrtKey  ", "") + _
               "Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= FrCoCd and Base.LocCd= FrToRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= FrCoCd and Loc.LocCd= FrFrRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and  ", "") + " FrToRmLoc<> '' " + _
               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') and Not (FrToRmLocTyp= Loc.LocTyp and FrToRmLocTyp= 'W') " + _
               wDtCndDetFg + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
               
  wSqlDetFgTo = wSqlDetFgTo1 + wSqlDetFgTo2
  moCn.Execute wSqlDetFgTo
    
'               "Where FrToRmLoc<> '' and substring(Base.LocTyp,1,1)<> 'X' and Base.LocTyp<> 'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') " + wDtCndDetFg + wCnd
  ' IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "
  wFgStr = " Union All " + wSqlOpnFgFr + " Union All " + wSqlOpnFgTo + _
           " Union All " + wSqlDetFgFr + " Union All " + wSqlDetFgTo
  
  '*** If Location 'W' is selected, an additional Unions FgRm FrLoc, FgRm ToLoc is required.
    
  
  
  '*** If Location 'Xs' or 'Fg' is selected, an additional Unions InRm FrLoc, InRm ToLoc is required.
  wSqlOpnInFr = "": wSqlOpnInTo = "": wSqlDetInFr = "": wSqlDetInTo = ""
  wFrDt1 = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "mm") + "/01", "01/" + Format(adc("UoDtFr"), "mm")) + "/" + Format(adc("UoDtFr"), "yy") '6.1
  
  '****** Sachin 3.02.0 - New Fields in Where Clause
  wDtCndOpnIn = " and InDt >= '" + wFrDt1 + "' And InDt < '" + CStr(adc("UoDtFr")) + "' and IrDt >= '" + wFrDt1 + "' And IrDt < '" + CStr(adc("UoDtFr")) + "' "

  adc("UoLn1Fr").CmpStr = "IrRmSz>= "
  adc("UoLn1To").CmpStr = "IrRmSz<= "
  adc("UoCstRtFr").CmpStr = "IrStkRt>= "
  adc("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "IrLotNo>= "
  adc("UoLotNoTo").CmpStr = "IrLotNo<= "
  adc("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' Zubin 212
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " IrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " IrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnInFr = "Insert Into #TmpStkSummLed Select " + grpflds5 + "," + _
                "'O' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, IrStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case IrFrRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case IrFrRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From InvRm NoLock Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey  ", "") + _
                "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey  ", "") + _
                "Join Loc Base On IrCoCd= LocCoCd and IrFrRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrFrRmLoc<> '' " + wDtCndOpnIn + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, IrLotNo, IrRmSz, IrRmSz2, IrRmSz3, IrStkRt Option (MaxDOP 1) "

'                "Where FrFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnFg + wCnd
' IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "
  
  moCn.Execute wSqlOpnInFr
  
  adc("UoLn1Fr").CmpStr = "IrRmSz>= "
  adc("UoLn1To").CmpStr = "IrRmSz<= "
  adc("UoCstRtFr").CmpStr = "IrStkRt>= "
  adc("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "IrLotNo>= "
  adc("UoLotNoTo").CmpStr = "IrLotNo<= "
  adc("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' Zubin 212
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " IrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " IrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****** (Jen 2.14) '' as qTc,'' as qYy, '' as qChr, 0 as qNo, " + _
  '                  "0 as qSr, 0 as qSrNo, '' as qBYy, '' as qBChr, 0 as qBNo, '' as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added and join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnInTo = "Insert Into #TmpStkSummLed Select " + grpflds5 + ", " + _
                "'O' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz  as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, IrStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case IrToRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case IrToRmDc When 'D' Then IrRmAWt When 'C' Then -IrRmAWt Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos " + _
                "From InvRm NoLock Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey  ", "") + _
                "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy " + _
                "and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey  ", "") + _
                "Join Loc Base On IrCoCd= LocCoCd and IrToRmLoc= LocCd " + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrToRmLoc<> '' " + wDtCndOpnIn + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, IrLotNo, IrRmSz, IrRmSz2, IrRmSz3, IrStkRt Option (MaxDOP 1) "

'                "Where FrToRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' " + _
'                "and Base.LocTyp<>'F' " + wDtCndOpnFg + wCnd
'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "
  moCn.Execute wSqlOpnInTo
  
  '****** Sachin 3.02.0 - New fields in Where Clause
  wDtCndDetIn = " and InDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' and IrDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"

  adc("UoLn1Fr").CmpStr = "IrRmSz>= "
  adc("UoLn1To").CmpStr = "IrRmSz<= "
  adc("UoCstRtFr").CmpStr = "IrStkRt>= "
  adc("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "IrLotNo>= "
  adc("UoLotNoTo").CmpStr = "IrLotNo<= "
  adc("UoLotNoSel").CmpStr = "IrLotNo In "
  
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' **** Zubin 211 **** '
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " IrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " IrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) IrTc as qTc, IrYy as qYy, IrChr as qChr, " + _
  '                  "IrNo as qNo, IrSr as qSr, IrSrNo as qSrNo, '' as qBYy, '' as qBChr, " + _
  '                  "0 as qBNo, InDt as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetInFr1 = "Insert Into #TmpStkSummLed Select " + grpflds6 + "," + _
               "'T' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, IrStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then IrRmAWt else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then IrRmAWt else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then IrRmAWt else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then IrRmAWt else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then IrRmAWt else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then IrRmAWt else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and IrFrRmDc='D' then IrRmAWt else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and IrFrRmDc='C' then IrRmAWt else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then IrRmAWt else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then IrRmAWt else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then IrRmAWt else 0 end) as qSal, " + _
               "(Case when IrFrRmDc='D' then 1 when IrFrRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then IrRmAWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetInFr2 = "From InvRm NoLock " + _
               "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey  ", "") + _
               "Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= IrCoCd and Base.LocCd= IrFrRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= IrCoCd and Loc.LocCd= IrToRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and  ", "") + " IrFrRmLoc<> '' " + wDtCndDetIn + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
  wSqlDetInFr = wSqlDetInFr1 + wSqlDetInFr2

  moCn.Execute wSqlDetInFr
  
'               "Where FrFrRmLoc<> '' and substring(Base.LocTyp,1,1)<>'X' and Base.LocTyp<>'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W')" + wDtCndDetFg + wCnd
'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "

  adc("UoLn1Fr").CmpStr = "IrRmSz>= "
  adc("UoLn1To").CmpStr = "IrRmSz<= "
  adc("UoCstRtFr").CmpStr = "IrStkRt>= "
  adc("UoCstRtTo").CmpStr = "IrStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "IrLotNo>= "
  adc("UoLotNoTo").CmpStr = "IrLotNo<= "
  adc("UoLotNoSel").CmpStr = "IrLotNo In "
  wrepcnd = adc.RepCond
  ' **** Zubin 211 **** '
  'wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  ' **** Zubin 211 **** '
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " IrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " IrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " IrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " IrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " IrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  ''''wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " IrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  ' Zubin 212
  '****** (Jen 2.14) IrTc as qTc, IrYy as qYy, IrChr as qChr, " + _
  '                  "IrNo as qNo, IrSr as qSr, IrSrNo as qSrNo, '' as qBYy, '' as qBChr, " + _
  '                  "0 as qBNo, InDt as qDt,
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  ' ***** Manali 3.7.1 - 'Xd' included in Cnv
  wSqlDetInTo1 = "Insert Into #TmpStkSummLed Select " + grpflds6 + "," + _
               "'T' as qFlag, RmCd, IrLotNo as qLotNo, IrRmSz as qRmSz, IrRmSz2 as qRmSz2, IrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, IrStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then IrRmAWt else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then IrRmAWt else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then IrRmAWt else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then IrRmAWt else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then IrRmAWt else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then IrRmAWt else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and IrToRmDc='D' then IrRmAWt else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and IrToRmDc='C' then IrRmAWt else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then IrRmAWt else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then IrRmAWt else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then IrRmAWt else 0 end) as qSal, " + _
               "(Case when IrToRmDc='D' then 1 when IrToRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then IrRmAWt else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
               "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 " + _
               "when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  ' **** Manali 3.03 - 09/07/08 - join with RmidSz added
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetInTo2 = "From InvRm NoLock " + _
               "join InvHd on InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IrPrtKey  ", "") + _
               "Join RmMst Rm On IrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=RmPrtKey  ", "") + _
               "Join Loc Base on Base.LocCoCd= IrCoCd and Base.LocCd= IrToRmLoc " + _
               "Join Loc Loc on Loc.LocCoCd= IrCoCd and Loc.LocCd= IrFrRmLoc " + _
               "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
               "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
               "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= IrRmSz and RrToLn>= IrRmSz " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
               "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and  ", "") + " IrToRmLoc<> '' " + wDtCndDetIn + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "
  wSqlDetInTo = wSqlDetInTo1 + wSqlDetInTo2
  
  moCn.Execute wSqlDetInTo
    
 '  IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "
 
  ' ***** Manali 3.7.1 - Stock Summary Changed for Trading
  wSqlOpnXdFr = "": wSqlOpnXdTo = "": wSqlDetXdFr = "": wSqlDetXdTo = ""
  wFrDt1 = IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "mm") + "/01", "01/" + Format(adc("UoDtFr"), "mm")) + "/" + Format(adc("UoDtFr"), "yy") '6.1
  
  wDtCndOpnXd = " and OmDt >= '" + wFrDt1 + "' And OmDt < '" + CStr(adc("UoDtFr")) + "' and OdOmDt >= '" + wFrDt1 + "' And OdOmDt < '" + CStr(adc("UoDtFr")) + "' " 'and BClsDt >= '" + wFrDt1 + "' And BClsDt < '" + CStr(adc("UoDtFr")) + "' "

  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  ' Zubin 212
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
 '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
  'wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
  'wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  
  Dim ws_FgRm As String
  ws_FgRm = " (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End) "
  
  wSqlOpnXdFr = "Insert Into #TmpStkSummLed Select " + grpflds7 + "," + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case OmFrRmDc When 'D' Then FrRmWt *" + ws_FgRm + " When 'C' Then -FrRmWt *" + ws_FgRm + " Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case OmFrRmDc When 'D' Then FrRmWt *" + ws_FgRm + " When 'C' Then -FrRmWt *" + ws_FgRm + " Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "

  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnXdFr = wSqlOpnXdFr + " From OrdMst NoLock " + _
                "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
                "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                "Join Loc Base On OmCoCd= LocCoCd and OmFrRmLoc= LocCd " + _
                "Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "' and ", "") + " OmFrRmLoc<> '' " + wDtCndOpnXd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " and (Select PMCd from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT')='JCN' " + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt Option (MaxDOP 1) "

'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "

  moCn.Execute wSqlOpnXdFr
    
  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  
  
  wSqlOpnXdTo = "Insert Into #TmpStkSummLed Select " + grpflds7 + ", " + _
                "'O' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescOpn + " as qRmSzDesc, RmCtg, Max(IsNull(PSCd,'')) As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt1 + " when RmCtg= 'X' then " + wXPureWtOrZwt1 + " else 1 end)" + wPurityWt1 + " as qRmPurityWt, " + _
                "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, " + _
                "sum(Case OmToRmDc When 'D' Then FrRmWt *" + ws_FgRm + " When 'C' Then -FrRmWt *" + ws_FgRm + " Else 0 End) as qOpn, " + _
                "0 as qPr, 0 as qFg, 0 as qCnv, 0 as qFLs, 0 as qWip, 0 as qS, " + _
                "0 as qPrdLsDr, 0 as qPrdLsCr, 0 as qBrk, 0 as qMsg, 0 as qSal, 1 as qMul, " + _
                "sum(Case OmToRmDc When 'D' Then FrRmWt *" + ws_FgRm + " When 'C' Then -FrRmWt *" + ws_FgRm + " Else 0 End) as qCls, " + _
                "max(Base.LocTyp) as qLocTyp, (Case max(Base.LocTyp) when 'S' then 1 " + _
                "when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 " + _
                "when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "

  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlOpnXdTo = wSqlOpnXdTo + " From OrdMst  NoLock " + _
                "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
                "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
                "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
                "Join Loc Base On OmCoCd= LocCoCd and OmToRmLoc= LocCd Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
                "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
                "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
                "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
                "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
                "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
                "Where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "' and ", "") + " OmToRmLoc<> '' " + wDtCndOpnXd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + _
                " and (Select PMCd from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT')='JCN' " + _
                " Group By Base.LocCoCd, Base.LocPrntCd, Base.LocCd, RmCtg, RmSCtg, RMCd, FrLotNo, FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt Option (MaxDOP 1) "

'                wDtCndOpnXd wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
'                IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then max(RmPurityWt) when RmCtg= 'X' then max(RmBasePurityWt) else 1 end) ", " 1 ") + wPurityWt1 + " as qRmPurityWt, "
  
  moCn.Execute wSqlOpnXdTo
  
  wDtCndDetXd = " and OmDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "' and OdOmDt between '" + CStr(adc("UoDtFr")) + "'" + _
              " And '" + CStr(adc("UoDtTo")) + "'"

  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrFrRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrFrRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  
  
  wSqlDetXdFr1 = "Insert Into #TmpStkSummLed Select " + grpflds8 + "," + _
               "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
               "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
               "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
               "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
               "(Case when Loc.LocTyp = 'Xp' then FrRmWt *" + ws_FgRm + " else 0 end) as qPr, " + _
               "(Case when Loc.LocTyp = 'F' then FrRmWt *" + ws_FgRm + " else 0 end) as qFg, " + _
               "(Case when Loc.LocTyp In ('Xc', 'Xd') then FrRmWt *" + ws_FgRm + " else 0 end) as qCnv, " + _
               "(Case when Loc.LocTyp = 'Xl' then FrRmWt *" + ws_FgRm + " else 0 end) as qFLs, " + _
               "(Case when Loc.LocTyp = 'W' then FrRmWt *" + ws_FgRm + " else 0 end) as qWip, " + _
               "(Case when Loc.LocTyp = 'S' then FrRmWt *" + ws_FgRm + " else 0 end) as qS, " + _
               "(Case when Loc.LocTyp = 'L' and OmFrRmDc='D' then FrRmWt *" + ws_FgRm + " else 0 end) as qPrdLsDr, " + _
               "(Case when Loc.LocTyp = 'L' and OmFrRmDc='C' then FrRmWt *" + ws_FgRm + " else 0 end) as qPrdLsCr, " + _
               "(Case when Loc.LocTyp = 'B' then FrRmWt *" + ws_FgRm + " else 0 end) as qBrk, " + _
               "(Case when Loc.LocTyp = 'M' then FrRmWt *" + ws_FgRm + " else 0 end) as qMsg, " + _
               "(Case when Loc.LocTyp = 'Xs' then FrRmWt *" + ws_FgRm + " else 0 end) as qSal, " + _
               "(Case   when OmFrRmDc='D' then 1 when OmFrRmDc='C' then -1 else 0 end) as qMul, " + _
               "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
               "then FrRmWt *" + ws_FgRm + "  else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
               "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 when 'Xc' then 8 When 'Xl' then 9 when 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
   
 'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
 wSqlDetXdFr2 = " From OrdMst  NoLock " + _
              "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
              "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
              "Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
              "Join Loc Base on Base.LocCoCd= OmCoCd and Base.LocCd= OmFrRmLoc " + _
              "Join Loc Loc on Loc.LocCoCd= OmCoCd and Loc.LocCd= OmToRmLoc " + _
              "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              "Where " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey='" + ctCurrPrtn + "' and  ", "") + " OmFrRmLoc<> '' " + _
              "and Not (Base.LocTyp= Loc.LocTyp) /*and Base.LocTyp= 'W')*/ " + _
              "and (Select PMCd from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT')='JCN' " + _
               wDtCndDetXd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "

              'wDtCndDetXd wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
              'IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
              'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "
              
  wSqlDetXdFr = wSqlDetXdFr1 + wSqlDetXdFr2
  moCn.Execute wSqlDetXdFr
  


  adc("UoLn1Fr").CmpStr = "FrRmSz>= "
  adc("UoLn1To").CmpStr = "FrRmSz<= "
  adc("UoCstRtFr").CmpStr = "FrRmStkRt>= "
  adc("UoCstRtTo").CmpStr = "FrRmStkRt<= "
  
  adc("UoLotNoFr").CmpStr = "FrLotNo>= "
  adc("UoLotNoTo").CmpStr = "FrLotNo<= "
  adc("UoLotNoSel").CmpStr = "FrLotNo In "
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  wAddlCnd = ""
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMCTGFR")) <> "", " FrRmCtg <= '" + Trim(adc("UORMCTGFR")) + "'", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGFR")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGFR")) <> "", " FrRmSCtg >= '" + Trim(adc("UORMSCTGFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGTO")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGTO")) <> "", " FrRmSCtg <= '" + Trim(adc("UORMSCTGTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UORMSCTGSEL")) <> "", " And ", "") + IIF(Trim(adc("UORMSCTGSEL")) <> "", " FrRmSCtg In (" + Trim(adc("UORMSCTGSEL")) + ")", "")
  
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPFR")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPFR")) <> "", " FrToRmLocTyp >= '" + Trim(adc("UOLOCTYPFR")) + "'", "")
'  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOLOCTYPTO")) <> "", " And ", "") + IIF(Trim(adc("UOLOCTYPTO")) <> "", " FrToRmLocTyp <= '" + Trim(adc("UOLOCTYPTO")) + "'", "")
  
  
  wSqlDetXdTo1 = "Insert Into #TmpStkSummLed Select " + grpflds8 + "," + _
                 "'T' as qFlag, RmCd, FrLotNo as qLotNo, FrRmSz as qRmSz, FrRmSz2 as qRmSz2, FrRmSz3 as qRmSz3, " + wRmSzDescDet + " as qRmSzDesc, RmCtg, IsNull(PSCd,'') As qStnCls, RmSCtg, FrRmStkRt as qRmStkRt, " + _
                 "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
                 "(Case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + wPureWtOrZWt + " when RmCtg= 'X' then " + wXPureWtOrZwt + " else 1 end)" + wPurityWt + " as qRmPurityWt, " + _
                 "Base.LocCoCd as qCoCd, Base.LocPrntCd as qLocPrntCd, Base.LocCd as qLoc, 0 as qOpn, " + _
                 "(Case when Loc.LocTyp = 'Xp' then FrRmWt *" + ws_FgRm + " else 0 end) as qPr, " + _
                 "(Case when Loc.LocTyp = 'F' then FrRmWt *" + ws_FgRm + " else 0 end) as qFg, " + _
                 "(Case when Loc.LocTyp In ('Xc', 'Xd') then FrRmWt *" + ws_FgRm + " else 0 end) as qCnv, " + _
                 "(Case when Loc.LocTyp = 'Xl' then FrRmWt *" + ws_FgRm + " else 0 end) as qFLs, " + _
                 "(Case when Loc.LocTyp = 'W' then FrRmWt *" + ws_FgRm + " else 0 end) as qWip, " + _
                 "(Case when Loc.LocTyp = 'S' then FrRmWt *" + ws_FgRm + " else 0 end) as qS, " + _
                 "(Case when Loc.LocTyp = 'L' and OmToRmDc='D' then FrRmWt *" + ws_FgRm + " else 0 end) as qPrdLsDr, " + _
                 "(Case when Loc.LocTyp = 'L' and OmToRmDc='C' then FrRmWt *" + ws_FgRm + " else 0 end) as qPrdLsCr, " + _
                 "(Case when Loc.LocTyp = 'B' then FrRmWt *" + ws_FgRm + " else 0 end) as qBrk, " + _
                 "(Case when Loc.LocTyp = 'M' then FrRmWt *" + ws_FgRm + " else 0 end) as qMsg, " + _
                 "(Case when Loc.LocTyp = 'Xs' then FrRmWt *" + ws_FgRm + " else 0 end) as qSal, " + _
                 "(Case when OmToRmDc='D' then 1 when OmToRmDc='C' then -1 else 0 end) as qMul, " + _
                 "(Case when Base.LocTyp In ('Xp', 'F', 'Xc', 'Xl', 'W', 'S', 'L', 'B', 'M', 'Xs', 'Xd') " + _
                 "then FrRmWt *" + ws_FgRm + "  else 0 end) as qCls, Base.LocTyp as qLocTyp, " + _
                 "(Case Base.LocTyp when 'S' then 1 when 'W' then 2 when 'L' then 3 " + _
                 "when 'B' then 4 when 'M' then 5 when 'F' then 6 when 'Xp' then 7 " + _
                 "when 'Xc' then 8 When 'Xl' then 9 When 'Xs' then 10 When 'Xd' Then 11 else 0 end) as qBaseLocTypPos "
  
  'JewelSmith.2 ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%' changed as PSCd = RmStnCls
  wSqlDetXdTo2 = " From OrdMst NoLock " + _
              "Join OrdDsg On  OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdOmIdNo=OmIdNo and OdPrtKey=OmPrtKey " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "join Fgd On FdCoCd=OdCoCd and FdBYy=OdBYy and FdBChr=OdBChr and FdBNo=OdBNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=OdPrtKey ", "") + _
              "join FgRm on FdIdNo=FrFdIdNo and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and " + _
              "FdChr= FrChr and FdNo= FrNo and FdSr= FrSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FrPrtKey ", "") + _
              "Join RmMst Rm On FrRmCd= RmCd " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=RmPrtKey ", "") + _
              "Join Loc Base on Base.LocCoCd= OmCoCd and Base.LocCd= OmToRmLoc " + _
              "Join Loc Loc on Loc.LocCoCd= OmCoCd and Loc.LocCd= OmFrRmLoc " + _
              "left outer join Param on PTyp= 'STNCLS' and PMCD=RmCtg and PSCd=RmStnCls " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
              "RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= FrRmSz and RrToLn>= FrRmSz " + _
              "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=FrRmSz " + _
              "Where " + IIF(gs_Partition = ctCurrPrtn, " FrPrtKey='" + ctCurrPrtn + "' and  ", "") + " OmToRmLoc<> '' " + _
              "and Not (Base.LocTyp= Loc.LocTyp) /*and Base.LocTyp= 'W')*/ and (Select PMCd from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT')='JCN' " + _
               wDtCndDetXd + wCnd + IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDOP 1) "


              'wDtCndDetXd wCnd + " and OdTc=(Select PMCd from Param Where PTyp='TC' and PMCd='JCN' and PValue='JT') " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + _
              'IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + " Option (MaxDop 1)"
            'IIF(adc("UoYN") = "Y", " (case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt else 1 end) ", " 1 ") + wPurityWt + " as qRmPurityWt, "
            
  wSqlDetXdTo = wSqlDetXdTo1 + wSqlDetXdTo2
  moCn.Execute wSqlDetXdTo
    
  ' ***** Manali 3.7.1 - Stock Summary Changed for Trading
  
  
'               "Where FrToRmLoc<> '' and substring(Base.LocTyp,1,1)<> 'X' and Base.LocTyp<> 'F' " + _
'               "and Not (Base.LocTyp= Loc.LocTyp and Base.LocTyp= 'W') " + wDtCndDetFg + wCnd

' **** Zubin 211 **** '
'''''
'''''  wInStr = " Union All " + wSqlOpnInFr + " Union All " + wSqlOpnInTo + _
'''''           " Union All " + wSqlDetInFr + " Union All " + wSqlDetInTo
'''''  '*** If Location 'Xs' or 'Fg' is selected, an additional Unions InvRm FrLoc, InvRm ToLoc is required.
'''''
'''''
'''''  wSqlStrg = wsqlOpnYyMm + " Union All " + wSqlOpnTdFr + " Union All " + wSqlOpnTdTo + _
'''''             " Union All " + wSqlDetFr + " Union All " + wSqlDetTo + wFgStr + wInStr
'''''
'''''  ADC.RepRecSource = wSqlStrg

  '*** (Bef 2.14)
  'adc.RepRecSource = "Select * From #TmpStkSummLed"
  '*** (Bef 2.14)
  ' **** Zubin 211 **** '

  '*** (Jen 2.14) grouping by qFlag is retained to hide all the opening records in the report designer
  ' **** Manali 3.03 - 09/07/08 - RmSzDesc added
  adc.RepRecSource = "Select max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
                     "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
                     "qFlag, RmCd, qLotNo, qRmSz, qRmSz2, qRmSz3, max(qRmSzDesc) as qRmSzDesc, RmCtg, max(qStnCls) As qStnCls, RmSCtg, qRmStkRt, " + _
                     "max(qPtrYN) as qPtrYN, qCoCd, qLocPrntCd, qLoc, sum(qOpn* qRmPurityWt) as qOpn, " + _
                     "sum(qPr* qMul* qRmPurityWt) as qPr, " + _
                     "sum(qFg* qMul* qRmPurityWt) as qFg, " + _
                     "sum(qCnv* qMul* qRmPurityWt) as qCnv, " + _
                     "sum(qFLs* qMul* qRmPurityWt) as qFLs, " + _
                     "sum(qWip* qMul* qRmPurityWt) as qWip, " + _
                     "sum(qS* qMul* qRmPurityWt) as qS, " + _
                     "sum(qPrdLsDr* qRmPurityWt) as qPrdLsDr, " + _
                     "sum(qPrdLsCr* qRmPurityWt) as qPrdLsCr, " + _
                     "sum(qBrk* qMul* qRmPurityWt) as qBrk, " + _
                     "sum(qMsg* qMul* qRmPurityWt) as qMsg, " + _
                     "sum(qSal* qMul* qRmPurityWt) as qSal, " + _
                     "sum(qCls* qMul* qRmPurityWt) as qCls, max(qLocTyp) as qLocTyp, " + _
                     "max(qBaseLocTypPos) as qBaseLocTypPos " + _
                     "From #TmpStkSummLed " + _
                     "Group By DGrp1, DGrp2, DGrp3, qFlag, qCoCd, qLocPrntCd, qLoc, RmCtg, RmSCtg, RmCd, qLotNo, qRmSz, qRmSz2, qRmSz3, qRmStkRt " + _
                     " having Cast(sum(qOpn) as Decimal(16,4))<> 0 or Cast(sum(qPr* qMul)  as Decimal(16,4))<> 0 or Cast(sum(qFg* qMul)  as Decimal(16,4))<> 0 or " + _
                     "   Cast(sum(qCnv* qMul)  as Decimal(16,4))<> 0 or Cast(sum(qFLs* qMul)   as Decimal(16,4))<> 0 or Cast(sum(qWip* qMul)   as Decimal(16,4))<> 0 or " + _
                     "    Cast(sum(qS* qMul)   as Decimal(16,4))<> 0 or Cast(sum(qPrdLsDr)   as Decimal(16,4))<> 0 or Cast(sum(qPrdLsCr)   as Decimal(16,4))<> 0 or Cast(sum(qBrk* qMul) as Decimal(16,4))<> 0 or " + _
                     "    Cast(sum(qMsg* qMul) as Decimal(16,4))<> 0 or Cast(sum(qSal* qMul)  as Decimal(16,4))<> 0 /* or Cast(sum(qCls* qMul)  as Decimal(16,4))<> 0 */ "
  '*** (Jen 2.14)

  Rep.TxtFrDt.SetText (adc("UoDtFr"))
  Rep.TxtToDt.SetText (adc("UoDtTo"))
  CRV_REP.DisplayGroupTree = False

  ' **** Zubin 211 **** '

  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    Rep.HLotNo.Suppress = True: Rep.FldLotNo.Suppress = True
  End If
  ' Zubin 212
  
  'KC.3- StkRt-when Show Stock Rate=YES then Stock Rate is visible in report
  'if it is NO then Stock Rate is invisible in report
 
  If adc("UoYN2") = "Y" Then
    Rep.hqRmStkRt.SetText "Stock Rt"
    Rep.FldRmStkRt.Suppress = False
  Else
    Rep.hqRmStkRt.SetText ""
    Rep.FldRmStkRt.Suppress = True
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 7)
  
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc
  
'  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                    "Else (If {rdo.qPtrYN}= 'Y' then ToText({rdo.qRmSz}, 4) Else ToText({rdo.qRmSz}, 3))"
'  GRP_REP.SetFormula Rep, "wRmSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} "
  
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

Exit Sub

ErrHndlr:
  moCn.Execute "Delete From #TmpStkSummLed"
  DispMsg Err.Description, etError
  adc.RepRecSource = "Select * From #TmpStkSummLed"
' **** Zubin 211 **** '
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
  'StkRt-To select this option(Show Stock rate) either Seo user can give their password or
  ' can enter SEOPWD password otherwise this option can't not select
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPureWt"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "ShowZWt"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2
    If .Value = Checked Then
       If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
            "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
            "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
            "Isnull ((Select PValue1 from Param where " + _
            "Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ), ''))") Then _
            .Value = Unchecked
    End If
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
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
  
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_RmCtgFor = True
  gb_RmZ = False
  'gs_DmTcTyp = "DM"
  'gs_TxnTcTyp = "BV"
  'gs_OmTcTyp = "SO"
  'gs_CmCtg = "C"
  'gs_RmGrp = "ORDDET"
  'gs_LabGrp = "ORDDET"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      LblPrntCd.Visible = True: adc("UOPRNTCDFR").Visible = True: adc("UOPRNTCDTO").Visible = True: adc("UOPRNTCDSEL").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      LblPrntCd.Visible = False: adc("UOPRNTCDFR").Visible = False: adc("UOPRNTCDTO").Visible = False: adc("UOPRNTCDSEL").Visible = False
  End If
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  ' Zubin 212
  If Not moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") Then
    LblLotNo.Visible = False: adc("UoLotNoFr").Visible = False
    adc("UoLotNoTo").Visible = False: adc("UoLotNoSel").Visible = False
    adc("UoLotNoFr") = "": adc("UoLotNoTo") = "": adc("UoLotNoSel") = ""
  End If
  ' Zubin 212
  Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)

  gb_CoCdFor = False '****** Sachin 2.14.0 - [Multi-Company]

  Select Case UCase(IdName)
  
  Case Is = UCase("UoIssLocFr"), UCase("UoPrntCdFr"), UCase("UoPrntCdSel")
    gs_LocTyp = "'" + adc("UoLocTypFr") + "'"
  Case Is = UCase("UoIssLocTo"), UCase("UoPrntCdTo")
    gs_LocTyp = "'" + adc("UoLocTypTo") + "'"
  Case Is = UCase("UoIssLocSel")
    gs_LocTyp = ""
  Case Is = UCase("UoBaseCdFr")
    If adc("UoRmCtgFr") <> UCase("X") Then
      Cancel = True: ErrMsg = "Can be Entered Only For Accessories": Exit Sub
    End If
  '*** Jay 3.01 **** [LotNo Err]
  Case Is = UCase("UoLotNoFr"), UCase("UoLotNoSel")
    gs_RmLocTyp = adc("UoLocTypFr")
    gs_Loc = adc("UoIssLocFr")
  Case Is = UCase("UoLotNoTo")
    gs_RmLocTyp = adc("UoLocTypTo")
    gs_Loc = adc("UoIssLocTo")
  '*** Jay 3.01 **** [LotNo Err]
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

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  '****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

  gs_LocTyp = "'" + adc("UoLocTypFr") + "'"
  
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  'StkRt-Set Show StockRate option is No
  adc("UoYN2") = "N"
  
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
  Rep.TxtCoCdFr.SetText adc("UOCOCDFR")
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
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

'StkRt-if password is change to blank, Show StockRate check box will turn into blank
Private Sub TxtPwd_Change()
    If Trim(TxtPwd.text) = "" Then
        ChkBoxArr(2).Value = Unchecked
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

