VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.Form EmrFrpRmReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Requirement"
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
      Left            =   11565
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5565
      TabIndex        =   105
      Top             =   9615
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   106
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
         TabIndex        =   107
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
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   60
      TabIndex        =   97
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
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6780
      TabIndex        =   104
      Top             =   9585
      Visible         =   0   'False
      Width           =   735
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   120
         TabIndex        =   101
         Top             =   30
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
      Height          =   9660
      Left            =   0
      TabIndex        =   100
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   99
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRmReq.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRmReq.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   103
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   98
               Top             =   360
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   102
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Based on RmQty"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   3
               Left            =   10240
               TabIndex        =   162
               Tag             =   "ShowBigPic"
               ToolTipText     =   "Based on Rm Quantity"
               Top             =   8340
               Width           =   1600
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Exact Size1"
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
               Left            =   5220
               TabIndex        =   161
               Tag             =   "ShowBigPic"
               ToolTipText     =   "Comparison on exact size1"
               Top             =   8340
               Width           =   1470
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Upto Range Size"
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
               Left            =   3120
               TabIndex        =   89
               Tag             =   "ShowBigPic"
               ToolTipText     =   "Check To Include Range size"
               Top             =   8340
               Width           =   1890
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Big Image"
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
               Height          =   240
               Index           =   4
               Left            =   6840
               TabIndex        =   92
               Tag             =   "ShowBigPic"
               ToolTipText     =   "Check To Show Big Picture"
               Top             =   2475
               Width           =   2250
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture "
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
               Left            =   6840
               TabIndex        =   91
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2160
               Width           =   2250
            End
            Begin VB.Frame FraOptPtr 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   11640
               TabIndex        =   150
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   7800
               Width           =   2385
               Begin VB.OptionButton OptPtr 
                  Caption         =   "Actual Ptr"
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
                  TabIndex        =   87
                  ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
                  Top             =   285
                  Width           =   1845
               End
               Begin VB.OptionButton OptPtr 
                  Caption         =   "Calculated Ptr"
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
                  TabIndex        =   86
                  ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
                  Top             =   0
                  Width           =   1905
               End
            End
            Begin VB.Frame FraOptRmVal 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   7950
               TabIndex        =   149
               ToolTipText     =   "Select Cost Value Or Sales Value Or No Value"
               Top             =   7800
               Width           =   1545
               Begin VB.OptionButton OptRmVal 
                  Caption         =   "No Value"
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
                  Left            =   0
                  TabIndex        =   84
                  ToolTipText     =   "Select Cost Value Or Sales Value Or No Value"
                  Top             =   570
                  Width           =   1365
               End
               Begin VB.OptionButton OptRmVal 
                  Caption         =   "Cost Value"
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
                  TabIndex        =   82
                  ToolTipText     =   "Select Cost Value Or Sales Value Or No Value"
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.OptionButton OptRmVal 
                  Caption         =   "Sales Value"
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
                  TabIndex        =   83
                  ToolTipText     =   "Select Cost Value Or Sales Value Or No Value"
                  Top             =   285
                  Width           =   1455
               End
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   1485
               TabIndex        =   148
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   7755
               Width           =   1485
               Begin VB.OptionButton OptRmPrd 
                  Caption         =   "Order Rm"
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
                  TabIndex        =   79
                  ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
                  Top             =   285
                  Width           =   1245
               End
               Begin VB.OptionButton OptRmPrd 
                  Caption         =   "Prd Rm"
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
                  TabIndex        =   78
                  ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
                  Top             =   0
                  Width           =   1245
               End
            End
            Begin VB.TextBox TxtPwd 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   4140
               MaxLength       =   30
               PasswordChar    =   "*"
               TabIndex        =   80
               Top             =   7755
               Width           =   2340
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4140
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1455
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "OrRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6840
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1460
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "OrRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4140
               TabIndex        =   4
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   580
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGTO"
               IdName          =   "UORMSCTGTO"
               CmpStr          =   "OrRmSCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6840
               TabIndex        =   5
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   580
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "OrRmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2265
               TabIndex        =   1
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   295
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "OrRmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1485
               TabIndex        =   3
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   580
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UORMSCTGFR"
               IdName          =   "UORMSCTGFR"
               Data            =   "ABCDW"
               CmpStr          =   "OrRmSCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1485
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1455
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "OrRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4140
               TabIndex        =   19
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   2030
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1TO"
               IdName          =   "UOLN1TO"
               CmpStr          =   "OrLn1 <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1485
               TabIndex        =   18
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   2030
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#0.000"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UOLN1FR"
               IdName          =   "UOLN1FR"
               CmpStr          =   "OrLn1 >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4140
               TabIndex        =   21
               ToolTipText     =   "Enter To Raw Material Breadth"
               Top             =   2315
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   5
               DataField       =   "UOLN2TO"
               IdName          =   "UOLN2TO"
               CmpStr          =   "OrLn2 <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1485
               TabIndex        =   20
               ToolTipText     =   "Enter From Raw Material Breadth"
               Top             =   2315
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   5
               DataField       =   "UOLN2FR"
               IdName          =   "UOLN2FR"
               CmpStr          =   "OrLn2 >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   4140
               TabIndex        =   23
               ToolTipText     =   "Enter To Sale Rate"
               Top             =   2600
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSALRTTO"
               IdName          =   "UOSALRTTO"
               CmpStr          =   "OrSalRt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   1485
               TabIndex        =   22
               ToolTipText     =   "Enter From Sale Rate"
               Top             =   2600
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOSALRTFR"
               IdName          =   "UOSALRTFR"
               CmpStr          =   "OrSalRt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   4140
               TabIndex        =   25
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   2880
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
               CmpStr          =   "OrCstRt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   1485
               TabIndex        =   24
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   2880
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
               CmpStr          =   "OrCstRt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   4140
               TabIndex        =   27
               ToolTipText     =   "Enter To Production Quantity"
               Top             =   3170
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPRDQTYTO"
               IdName          =   "UOPRDQTYTO"
               CmpStr          =   "OdPrdQty <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   1485
               TabIndex        =   26
               ToolTipText     =   "Enter From Production Quantity"
               Top             =   3170
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOPRDQTYFR"
               IdName          =   "UOPRDQTYFR"
               CmpStr          =   "OdPrdQty >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   4140
               TabIndex        =   29
               ToolTipText     =   "Enter To Balance Production"
               Top             =   3455
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDTO"
               IdName          =   "UOBALPRDTO"
               CmpStr          =   "(OdPrdQty-  OdFgQty) <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1485
               TabIndex        =   28
               ToolTipText     =   "Enter From Balance Production"
               Top             =   3455
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDFR"
               IdName          =   "UOBALPRDFR"
               CmpStr          =   "(OdPrdQty-  OdFgQty)>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4140
               TabIndex        =   42
               ToolTipText     =   "Enter To Design Category"
               Top             =   4310
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6840
               TabIndex        =   43
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   4310
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1485
               TabIndex        =   41
               ToolTipText     =   "Enter From Design Category"
               Top             =   4310
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4140
               TabIndex        =   45
               ToolTipText     =   "Enter To Design Code"
               Top             =   4595
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "OdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6840
               TabIndex        =   46
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4595
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1485
               TabIndex        =   44
               ToolTipText     =   "Enter From Design Code"
               Top             =   4595
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "OdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4140
               TabIndex        =   49
               ToolTipText     =   "Enter To Design Size"
               Top             =   4880
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "OdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1485
               TabIndex        =   47
               ToolTipText     =   "Enter From Design Size"
               Top             =   4880
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "OdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4800
               TabIndex        =   50
               ToolTipText     =   "Enter To Suffix"
               Top             =   4880
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "OdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2130
               TabIndex        =   48
               ToolTipText     =   "Enter From Suffix"
               Top             =   4880
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "OdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4140
               TabIndex        =   52
               ToolTipText     =   "Enter To Customer"
               Top             =   5165
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6840
               TabIndex        =   53
               ToolTipText     =   "Enter Customer Selection"
               Top             =   5165
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1485
               TabIndex        =   51
               ToolTipText     =   "Enter From Customer"
               Top             =   5165
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4140
               TabIndex        =   55
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   5450
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
               Index           =   44
               Left            =   6840
               TabIndex        =   56
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   5450
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1485
               TabIndex        =   54
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   5450
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
               Index           =   46
               Left            =   4140
               TabIndex        =   66
               ToolTipText     =   "Enter To Order Date"
               Top             =   6305
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
               Left            =   1485
               TabIndex        =   65
               ToolTipText     =   "Enter From Order Date"
               Top             =   6305
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
               Index           =   49
               Left            =   4140
               TabIndex        =   68
               ToolTipText     =   "Enter To Prd Delivery Date"
               Top             =   6605
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "OdDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1485
               TabIndex        =   67
               ToolTipText     =   "Enter From Prd Delivery Date"
               Top             =   6605
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "OdDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4905
               TabIndex        =   34
               ToolTipText     =   "Enter To Order Year"
               Top             =   3740
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
               Left            =   2265
               TabIndex        =   31
               ToolTipText     =   "Enter From Order Year"
               Top             =   3740
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
               Left            =   1485
               TabIndex        =   30
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3740
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
               Left            =   2685
               TabIndex        =   32
               ToolTipText     =   "Enter From Order Character"
               Top             =   3740
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
               Left            =   4140
               TabIndex        =   33
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3740
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
               Left            =   5325
               TabIndex        =   35
               ToolTipText     =   "Enter To Order Character"
               Top             =   3740
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
               Index           =   26
               Left            =   5025
               TabIndex        =   39
               ToolTipText     =   "Enter To Order Serial"
               Top             =   4025
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
               Index           =   52
               Left            =   2385
               TabIndex        =   37
               ToolTipText     =   "Enter From Order Serial"
               Top             =   4025
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
               Index           =   53
               Left            =   1485
               TabIndex        =   36
               ToolTipText     =   "Enter From Order Number"
               Top             =   4025
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
               Left            =   4140
               TabIndex        =   38
               ToolTipText     =   "Enter To Order Number"
               Top             =   4025
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
               Index           =   25
               Left            =   4140
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Sub Shape To Range"
               Top             =   870
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSUBSHPTO"
               IdName          =   "UOSUBSHPTO"
               CmpStr          =   "OrSubShp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   6840
               TabIndex        =   8
               ToolTipText     =   "Enter Raw Material Sub Shape Selection"
               Top             =   870
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSUBSHPSEL"
               IdName          =   "UOSUBSHPSEL"
               CmpStr          =   "OrSubShp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1485
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Sub Shape From Range"
               Top             =   870
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSUBSHPFR"
               IdName          =   "UOSUBSHPFR"
               CmpStr          =   "OrSubShp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   1485
               TabIndex        =   77
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   7755
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOYN1,UOYN"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1485
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   295
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7950
               TabIndex        =   81
               ToolTipText     =   "Enter 'Y' For Cost Value, 'N' for Sale Value or Blank For No Value."
               Top             =   7800
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   1485
               TabIndex        =   71
               ToolTipText     =   "Enter From  Priority Code"
               Top             =   7175
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
               CmpStr          =   "OdPrtCd>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   4140
               TabIndex        =   72
               ToolTipText     =   "Enter To  Priority Code"
               Top             =   7175
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
               CmpStr          =   "OdPrtCd<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   6840
               TabIndex        =   73
               ToolTipText     =   "Enter  Priority Code Selection"
               Top             =   7175
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd  In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   11640
               TabIndex        =   85
               ToolTipText     =   "Enter 'Y' For Calculated Pointer, 'N' for Stored Actual Pointer"
               Top             =   7800
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN1"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   1485
               TabIndex        =   15
               ToolTipText     =   "Enter From Suffix"
               Top             =   1745
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZFR"
               IdName          =   "UORNGSZFR"
               CmpStr          =   "RrRngSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   4140
               TabIndex        =   16
               ToolTipText     =   "Enter From Suffix"
               Top             =   1745
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
               CmpStr          =   "RrRngSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   6840
               TabIndex        =   17
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1745
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORNGSZSEL"
               IdName          =   "UORNGSZSEL"
               CmpStr          =   "RrRngSz In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   1485
               TabIndex        =   88
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   8340
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOCURCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   9255
               TabIndex        =   90
               ToolTipText     =   "Enter 'Y' For Picture"
               Top             =   2160
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN2"
               ReCalcParent    =   "UOYN2"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   2775
               TabIndex        =   58
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   5730
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   5430
               TabIndex        =   60
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   5730
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   66
               Left            =   6840
               TabIndex        =   61
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   5735
               Visible         =   0   'False
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "OdPoNo in"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   1485
               TabIndex        =   93
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   8650
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               ReCalcParent    =   "UOBCHRFR"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   68
               Left            =   3120
               TabIndex        =   94
               ToolTipText     =   "Enter Bag Character"
               Top             =   8650
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOBCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   4140
               TabIndex        =   2
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   295
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "OrRmCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   4140
               TabIndex        =   10
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1160
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
               Index           =   71
               Left            =   1485
               TabIndex        =   9
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1160
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
               Index           =   72
               Left            =   6840
               TabIndex        =   11
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1160
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPARENTCDSEL"
               IdName          =   "UOPARENTCDSEL"
               CmpStr          =   "RmParentCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   4140
               TabIndex        =   70
               ToolTipText     =   "Enter To Exp Delivery Date"
               Top             =   6890
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1TO"
               IdName          =   "UODELDT1TO"
               CmpStr          =   "OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   1485
               TabIndex        =   69
               ToolTipText     =   "Enter From Exp Delivery Date"
               Top             =   6890
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDT1FR"
               IdName          =   "UODELDT1FR"
               CmpStr          =   "OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   9225
               TabIndex        =   155
               ToolTipText     =   "Enter 'Y' For Big Picture"
               Top             =   2460
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   4140
               TabIndex        =   63
               ToolTipText     =   "Enter To Production Sequence"
               Top             =   6020
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "OdPrdSeq <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   6840
               TabIndex        =   64
               ToolTipText     =   "Enter Production Sequence Selection"
               Top             =   6020
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPMCDSEL"
               IdName          =   "UOPMCDSEL"
               CmpStr          =   "OdPrdSeq In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   78
               Left            =   1485
               TabIndex        =   62
               ToolTipText     =   "Enter From Productiopn Sequence"
               Top             =   6020
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "OdPrdSeq >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   79
               Left            =   1485
               TabIndex        =   57
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   5730
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
               CmpStr          =   "OdSubCust >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   4140
               TabIndex        =   59
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   5730
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "OdSubCust <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   81
               Left            =   5520
               TabIndex        =   160
               ToolTipText     =   "Enter 'Y' to Include Range Size"
               Top             =   8340
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   6840
               TabIndex        =   40
               ToolTipText     =   "Enter Order Number Selection "
               Top             =   4025
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOOMNOSEL"
               IdName          =   "UOOMNOSEL"
               CmpStr          =   "OdNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   88
               Left            =   1485
               TabIndex        =   74
               ToolTipText     =   "Enter From Raw Material Stone Class "
               Top             =   7460
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "ISNULL(SC.PSCd,'') >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   89
               Left            =   4140
               TabIndex        =   75
               ToolTipText     =   "Enter To Raw Material Stone Class "
               Top             =   7460
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "ISNULL(SC.PSCd,'') <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   90
               Left            =   6840
               TabIndex        =   76
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   7460
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
               Index           =   43
               Left            =   60
               TabIndex        =   164
               Top             =   7440
               Width           =   1305
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
               Index           =   42
               Left            =   5370
               TabIndex        =   159
               Top             =   5730
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
               Index           =   41
               Left            =   2715
               TabIndex        =   158
               Top             =   5730
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Sequence"
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
               Index           =   40
               Left            =   60
               TabIndex        =   157
               Top             =   6020
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "For Per Pc detail take Group Sort as Detail"
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
               Height          =   450
               Index           =   39
               Left            =   5160
               TabIndex        =   156
               Top             =   8700
               Width           =   9615
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Del Dt"
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
               Index           =   38
               Left            =   60
               TabIndex        =   154
               Top             =   6890
               Width           =   1425
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
               Index           =   37
               Left            =   60
               TabIndex        =   153
               Top             =   1160
               Width           =   1425
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Type"
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
               TabIndex        =   152
               ToolTipText     =   "Location"
               Top             =   8650
               Width           =   1035
            End
            Begin VB.Label LblUoBChrFr 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Chr"
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
               Left            =   2280
               TabIndex        =   151
               Top             =   8650
               Width           =   885
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sub PO No"
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
               Index           =   36
               Left            =   60
               TabIndex        =   147
               Top             =   5735
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Picture "
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
               Height          =   270
               Index           =   35
               Left            =   6910
               TabIndex        =   146
               Top             =   2160
               Visible         =   0   'False
               Width           =   1335
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
               TabIndex        =   145
               Top             =   8340
               Width           =   1005
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
               Index           =   34
               Left            =   4740
               TabIndex        =   144
               Top             =   4880
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
               Index           =   33
               Left            =   2070
               TabIndex        =   143
               Top             =   4880
               Width           =   105
            End
            Begin VB.Label ALBL 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Show Pointer"
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
               Height          =   450
               Index           =   32
               Left            =   9810
               TabIndex        =   142
               Top             =   7800
               Width           =   1755
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Priority Cd"
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
               Height          =   255
               Left            =   60
               TabIndex        =   141
               Top             =   7175
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Enter SEO Password"
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
               Height          =   450
               Index           =   29
               Left            =   3150
               TabIndex        =   140
               Top             =   7755
               Width           =   975
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Value"
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
               Height          =   450
               Index           =   28
               Left            =   6900
               TabIndex        =   139
               Top             =   7800
               Width           =   945
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
               Index           =   23
               Left            =   2205
               TabIndex        =   138
               Top             =   295
               Width           =   105
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Prd Rm Qty && Wt"
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
               Height          =   525
               Left            =   60
               TabIndex        =   137
               Top             =   7725
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Shp"
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
               TabIndex        =   136
               Top             =   870
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
               Left            =   60
               TabIndex        =   135
               Top             =   3740
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
               Left            =   60
               TabIndex        =   134
               Top             =   4025
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
               Left            =   5265
               TabIndex        =   133
               Top             =   3740
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
               Left            =   2625
               TabIndex        =   132
               Top             =   3740
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
               Left            =   4845
               TabIndex        =   131
               Top             =   3740
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
               Left            =   2205
               TabIndex        =   130
               Top             =   3740
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
               Left            =   4965
               TabIndex        =   129
               Top             =   4025
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
               Left            =   2325
               TabIndex        =   128
               Top             =   4025
               Width           =   105
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   20
               Left            =   60
               TabIndex        =   127
               Top             =   6605
               Width           =   1425
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
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   19
               Left            =   60
               TabIndex        =   126
               Top             =   6305
               Width           =   1425
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
               Left            =   60
               TabIndex        =   125
               Top             =   5450
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
               TabIndex        =   124
               Top             =   5165
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Range Sz"
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
               TabIndex        =   123
               Top             =   1745
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Sz/Sfx"
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
               Left            =   60
               TabIndex        =   122
               Top             =   4880
               Width           =   1425
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
               TabIndex        =   121
               Top             =   4595
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Ctg"
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
               TabIndex        =   120
               Top             =   4310
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Bal Prd"
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
               TabIndex        =   119
               Top             =   3455
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Qty"
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
               TabIndex        =   118
               Top             =   3170
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cost Rate"
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
               TabIndex        =   117
               Top             =   2880
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sale Rate"
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
               TabIndex        =   116
               Top             =   2600
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Breadth"
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
               TabIndex        =   115
               Top             =   2315
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Len/Sv"
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
               TabIndex        =   114
               Top             =   2030
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Co Cd/Rm Ctg"
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
               TabIndex        =   113
               ToolTipText     =   "Location"
               Top             =   295
               Width           =   1485
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
               Left            =   60
               TabIndex        =   112
               Top             =   580
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
               TabIndex        =   111
               Top             =   1460
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
               TabIndex        =   110
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
               Left            =   4140
               TabIndex        =   109
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
               Left            =   1485
               TabIndex        =   108
               Top             =   0
               Width           =   1335
            End
         End
      End
   End
   Begin VB.OLE OLE1 
      Class           =   "Paint.Picture"
      Height          =   195
      Left            =   2430
      OleObjectBlob   =   "EmrFrpRmReq.frx":0038
      TabIndex        =   163
      Top             =   9780
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "EmrFrpRmReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRmReq
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mf_CalcPtr As Boolean
Dim mb_ExactSz1 As Boolean
Dim mb_OnQty As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyH And Shift = 2 Then
  KeyCode = 0
  Dim wPath As String
  On Error GoTo ErrH
  If MsgBox("This will open the Help document for this Menu in a separate window, Proceed?", vbYesNo, "Open Help File") = vbYes Then
  wPath = adc.Connection.GetFldVal("Select PDesc225 from Param Where PTyp='PCT' and PMCd='3D'") + "\Help\" + adc.MenuCd + ".doc"
  OLE1.CreateLink wPath
  OLE1.DoVerb
  Exit Sub
  Else
    Exit Sub
  End If
ErrH:
  If Err.Number = 31031 Then
    MsgBox "File/Folder (" + wPath + ") Not Found", etError
  Else
    DispMsg Err.Description, etError
  End If
  End If
End Sub



Private Sub SetGroupSort()
  GRP_REP.Add "Rm Code", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)+Rtrim(OrRmCd)", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hOrRmCtg,hOrRmSCtg,hOrRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "
  GRP_REP.Add "Rm Ctg", "Rtrim(OrRmCtg)", "", "OrRmCtg", "hOrRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)", "'('+ RTrim(OrRmCtg)+ ') '+ OrRmSCtg", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) "
  GRP_REP.Add "Sub Shp", "OrSubShp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'SUBSHP' and PMCd = OrSubShp) "
  '6.1-98 Stone Class group sort added
  GRP_REP.Add "Stone Class", "RmCtg+IsNull(SC.PSCd, '')", "'('+ RTrim(RmCtg)+') '+IsNull(SC.PSCd, '')", "OrRmCtg", "hOrRmCtg", "Rm Ctg", "", "", "SC.PDesc"
 
  '4.1.2
  GRP_REP.Add "Parent Cd", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)+(case when RmParentCd= '' then OrRmCd else RmParentCd end)", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ (case when RmParentCd= '' then OrRmCd else RmParentCd end)", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  '###
  'GRP_REP.Add "Rm Len/Sv", "Str(OrLn1,6,3)", "", "OrLn1", "hOrLn1", "", "", "", ""
  'GRP_REP.Add "Rm Len 2", "str(OrLn1,6,3)+' '+str(OrLn1,6,3)", "'('+ LTrim(str(OrLn1,6,3))+ ') '+ LTrim(str(OrLn2,6,3))", "OrLn1,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
 
  'GRP_REP.Add "Rm Len/Sv", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else Str(OrLn1,6,3) end)", "", "wLnSv", "hOrLn1", "", "", "", ""
   '****************Geeta**************Emr206*********12/04/04
   Dim wsPtrYN As String
   wsPtrYN = "PValue3 "
'   GRP_REP.Add "Rm Len/Sv", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else (case when " + wsPtrYN + "='Y' then Str(OrLn1,7,4) else Str(OrLn1,6,3) end) end)", "", "wLnSv", "hOrLn1", "", "", "", ""
'   GRP_REP.Add "Rm Len 2", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else Str(OrLn1,6,3) end)+' '+str(OrLn2,6,3)", "'('+ (case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else LTrim(Str(OrLn1,6,3)) end)+ ') '+ LTrim(str(OrLn2,6,3))", "wLnSv,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
'   GRP_REP.Add "Rm Sz", " str(Orln1,8,4 )+STR(OrLn2,8,4)+STR(OrLn3,8,4)", "(case when OrLn1+OrLn2+OrLn3=0 then '' else (str(Orln1,8,4 )+' * '+STR(OrLn2,8,4)+' * '+STR(OrLn3,8,4)) end)", "wRmSz", "hRmsz", "Rm Ctg,Rm SubCtg,Rm Cd,Range Size,Rm Ptr", "", "", ""
    
    GRP_REP.Add "Rm Sz", GetRmSzDBFldNm("OrLn1", "OrLn2", "OrLn3"), GetRmSzDispFldNm("OrLn1", "OrLn2", "OrLn3"), "wRmSz", "hRmsz", "Rm Ctg,Rm SubCtg,Rm Cd,Range Size,Rm Ptr", "", "", ""

'    Call GetRmSzGrpSrt("OrLn1", "OrLn2", "OrLn3", GRP_REP)
    ' ****** sachin 2.12 - 24/11/05 ****** '
    GRP_REP.Add "Range Size", "RrRngSz", "", "RrRngSz", "hRrRngSz", "", "", "", ""
    ' ****** sachin 2.12 - 24/11/05 ****** '
  
  'Manoj 2.10.0
  'If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", "Str(OrRmPtr,7,4)", "", "OrRmPtr", "hOrRmPtr", "", "", "", ""
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", Array("Str(OrRmPtr,7,4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End,7,4)"), Array("Str(OrRmPtr, 7, 4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End, 7, 4)"), "OrRmPtr", "hOrRmPtr", "", "", "", ""
  'Manoj 2.10.0

  ' Zubin 213 (Bef 213)
  'GRP_REP.Add "Rm Sale Rate", "Str(OrSalRt, 10, 3)", "", "OrSalRt", "hOrSalRt", "RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
  'GRP_REP.Add "Rm Cost Rate", "str(OrCstRt, 10, 3)", "", "OrCstRt", "hOrCstRt", "RmCd Sale Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
  'GRP_REP.Add "(RmCd) SaleRt", "RmCd+str(OrSalRt, 10, 3)", "'('+RmCd+') '+LTrim(str(OrSalRt,9,2))", "OrRmCtg,OrRmSCtg,OrRmCd,OrSalRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrSalRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) CostRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  'GRP_REP.Add "(RmCd) CostRt", "RmCd+str(OrCstRt, 10, 3)", "'('+RmCd+') '+LTrim(str(OrCstRt,9,2))", "OrRmCtg,OrRmSCtg,OrRmCd,OrCstRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrCstRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  
  ' Zubin 213
  GRP_REP.Add "Rm Sale Rate", "'['+RmQw+']'+Str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) else OrSalRt end, 10, 3)", _
          "", "OrSalRt", "hOrSalRt", "RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
  GRP_REP.Add "Rm Cost Rate", "'['+RmQwCst+']'+str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) else OrCstRt end, 10, 3)", _
          "", "OrCstRt", "hOrCstRt", "RmCd Sale Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
  GRP_REP.Add "(RmCd) SaleRt", "RmCd+str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) else OrSalRt end, 10, 3)", _
          "'('+RmCd+') '+LTrim(str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) else OrSalRt end,9,2))", _
          "OrRmCtg,OrRmSCtg,OrRmCd,OrSalRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrSalRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) CostRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  GRP_REP.Add "(RmCd) CostRt", "RmCd+str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) else OrCstRt end, 10, 3)", _
          "'('+RmCd+') '+LTrim(str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) else OrCstRt end,9,2))", _
          "OrRmCtg,OrRmSCtg,OrRmCd,OrCstRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrCstRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  ' Zubin 213

  'GRP_REP.Add "RmCd Sale Rate", "OrSalRt", "", "OrSalRt", "hOrSalRt", "OrCstRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "
  'GRP_REP.Add "RmCd Cost Rate", "OrCstRt", "", "OrCstRt", "hOrCstRt", "OrSalRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "
  '###
  
  'GRP_REP.Add "(RmCd) Len/Sv", "OrRmCd+Str(OrLn1,4,2)", "'('+OrRmCd+') '+LTrim(str(OrLn1,4,2))", "OrRmCd,OrLn1", "hOrRmCd,hOrLn1", "Rm Ctg,Rm SubCtg,Rm Code", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  'GRP_REP.Add "(RmCd) Len 2", "OrRmCd+str(OrLn1,4,2)+' '+str(OrLn1,4,2)", "OrRmCd+' ('+ LTrim(str(OrLn1,4,2))+ ') '+ LTrim(str(OrLn2,4,2))", "OrRmCd,OrLn1,OrLn2", "hOrRmCd,hOrLn1,hOrLn2", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Len/Sv", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  
  'GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd,OmDt,OmPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo", "Order Dt,Customer Code,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
  '6.1-6
  GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd", "hOrdNo,hOmCmCd", "Order Dt,Customer Code,Customer PO", "", "", "'('+convert(varchar(8), OmDt, " + CStr(DtFmtDashVal()) + ") +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
  GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,(DsgCd)Colour", "", "", "(Select DmDesc From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '') "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'PRTCD' and PmCd= OdPrtCd) "
  '6.1-6
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "wOrdNo,OmCmCd,OmDt,OmPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo", "Order Dt,Customer Code,OrderNo", "", "", "'('+convert(varchar(8), OmDt, " + CStr(DtFmtDashVal()) + ") +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
  '*** Jay 3.2.0 [OdPoNo]
  'vk.19 sub customer added
  '6.1-6
  GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' or OdSubCust<>'' then '/'+ LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )", "wOrdNo,OmCmCd,OmDt,OmPoNo,OdPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo,hOdPoNo", "Order Dt,Customer Code,OrderNo,Customer PO", "", "", "'('+convert(varchar(8), OmDt, " + CStr(DtFmtDashVal()) + ") +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))"
  '*** Jay 3.2.0 [OdPoNo]
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "OdDmCd", "hOdDmCd", "Design Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "(OrdKt)Colour", "OdKt+OdDmCol", "'('+OdKt+') '+OdDmCol", "OdKt,OdDmCol", "", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt)+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
' **** Zubin 211 **** '
  '6.1-6
  GRP_REP.Add "OrderSr", "OdTc+OdYy+OdChr+str(OdNo)+str(OdSr)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))", "wOrdNo,OmCmCd,OrSr", "hOrdNo,hOmCmCd,hSr", "OrderNo,Order Dt,Customer Code,Customer PO,Design Ctg,Design Code,Priority,(DsgCd)Colour,(OrdKt)Colour", "", "", "'('+convert(varchar(8), OmDt, " + CStr(DtFmtDashVal()) + ") +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
' **** Zubin 211 **** '

  'Sachin 4.1.2
  '6.1-6
  GRP_REP.Add "Exp Del Date", "convert(varchar(10), OdExpDelDt, 112)", "convert(varchar(8), OdExpDelDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
  
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'****************Geeta**********Emr206*******qPtrYN Included
'*** Jay 3.2.0 [PicNm]
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  OrRmCd, OrRmCtg, OrRmSCtg, PValue3 as qPtrYN, OrLn1, OrLn2, OrRmPtr, OrQty,
'  OrWt as qOrCarats,  OrWt as qOrWt, OrWt as qOrPureWt,
'  OrSalRt, OrCstRt, OdPrdQty, OdPrdQty as qOdBalPrd,
'  OrTc, OrYy,OrChr,OrNo,OrSr, OrSrNo, OdDmCd, OdDmSz, OdSfx, OdPicNm,
'  OmCmCd , OmPoNo, DmTcTyp, DmCtg, OmDt, OdDelDt
'  from OrdRm, OrdDsg, OrdMst, DsgMst,Param  where 1=2
  '*** Report Sql ***

' *************************************** Zubin 211 *************************************** '
' ***** PN: Whenever SQL of Rm Requirement (Static) changes, change the SQL of Rm Stock Req report also ***** '
' *************************************** Zubin 211 *************************************** '

  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wPurityWt As String
  Dim ws_RmQty As String, ws_RmWt As String, w_RsRmVal As String, ws_RmPtr As String
  
  '****** Sachin 2.14.0 - [Multi-Company] *****
  Dim wGrp(6) As String, wDetPos As Integer, i As Integer
  '****** Sachin 2.14.0 - [Multi-Company] *****
  
  Dim wRtSrtPos As Integer     ' ******* Manali 3.5.0 - 11/11/08 - Rate Group Sorts Locked SEO Pwd
  
  ' **** Zubin 211 **** '
  Dim ws_OrdBy As String
  ' **** Zubin 211 **** '

  ' Zubin 213
  Dim ws_Join As String
  Dim wb_BaseCur As Boolean
  Dim wAddlCnd As String  '****** Sachin 3.02.0 - Additional New Fields
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
      
  ' **** Crm 2.12 - 23-02-06 **** '
  Dim wCrmCnd As String
  wCrmCnd = " And OmSubmitYN <> 'N' "
  ' **** Crm 2.12 - 23-02-06 **** '
  
  If adc("UoYN3") = "Y" Then
    ws_RmQty = "OrPrdQty as OrQty"
    ws_RmWt = "OrPrdWt"
  Else
    ws_RmQty = "OrQty"
    ws_RmWt = "OrWt"
  End If
  
  
  
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)

  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
        wDetPos = i: Exit For
    End If
  Next i
  '****** Sachin 2.14.0 - [Multi-Company] *****
  
  ' ******* Manali 3.5.0 - 11/11/08 - Rate Group Sorts Locked SEO Pwd
  i = 1
  For i = 1 To 3
    If wGrp(i) = "Rm Sale Rate" Or wGrp(i) = "Rm Cost Rate" Or wGrp(i) = "(RmCd) SaleRt" Or wGrp(i) = "(RmCd) SaleRt" Then
      wRtSrtPos = i: Exit For
    End If
  Next i
 If ChkBoxArr(0).Value = vbChecked And ChkBoxArr(1).Value = vbChecked Then
  DispMsg "Please select either Range Size or Size1", etError
  Exit Sub
 End If
 If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") And wRtSrtPos <> 0 Then _
  DispMsg "Rate Group Sorts Can be Selected If SEO Password is Entered", etError: Exit Sub
  
 If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") And wRtSrtPos <> 0 Then _
  DispMsg "Rate Group Sorts Can be Selected If SEO Password is Entered", etError: Exit Sub
  ' ******* Manali 3.5.0 - 11/11/08 - Rate Group Sorts Locked SEO Pwd
  
  If adc("UoYN1") = "Y" And adc("UoYN3") = "Y" Then
    ws_RmPtr = " Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End As OrRmPtr "
    grpflds = GRP_REP.GrpFldLst(1)
  Else
    ws_RmPtr = " OrRmPtr "
    grpflds = GRP_REP.GrpFldLst(0)
  End If
  'Manoj 2.10.0
  
  If adc("UoYn") = "N" Then
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qOdBalPrd}*{rdo.OrSalRt}"      '"{rdo.OdPrdQty}*{rdo.OrSalRt}"
    GRP_REP.SetFormula Rep, "g1_wValDisp", "'Rm SalVal: ' + ToText ({@g1_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g2_wValDisp", "'Rm SalVal: ' + ToText ({@g2_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g3_wValDisp", "'Rm SalVal: ' + ToText ({@g3_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g_wValDisp", "'Rm SalVal: ' + ToText ({@g_wRmVal}, 2)"
  ElseIf adc("UOYN") = "Y" Then
    GRP_REP.SetFormula Rep, "wRmVal", "{rdo.qOdBalPrd}*{rdo.OrCstRt}"      '"{rdo.OdPrdQty}*{rdo.OrCstRt}"
    GRP_REP.SetFormula Rep, "g1_wValDisp", "'Rm CstVal: ' + ToText ({@g1_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g2_wValDisp", "'Rm CstVal: ' + ToText ({@g2_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g3_wValDisp", "'Rm CstVal: ' + ToText ({@g3_wRmVal}, 2)"
    GRP_REP.SetFormula Rep, "g_wValDisp", "'Rm CstVal: ' + ToText ({@g_wRmVal}, 2)"
  ElseIf adc("UOYN") = "" Then
    GRP_REP.SetFormula Rep, "wRmVal", "0"
    GRP_REP.SetFormula Rep, "g1_wValDisp", ""
    GRP_REP.SetFormula Rep, "g2_wValDisp", ""
    GRP_REP.SetFormula Rep, "g3_wValDisp", ""
    GRP_REP.SetFormula Rep, "g_wValDisp", ""
  End If
'  GRP_REP.SetFormula Rep, "wRmSzDisp", "GetRmSzDisp1(FldRmSz1.Value, FldRmSz2.Value, FldRmSz3.Value)"
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  ' Zubin 213
  wb_BaseCur = IIF(UCase(adc("UoCurCdFr")) = "", True, False)
  wCnd = wCnd + IIF(wb_BaseCur = False, " and OmCmCurCd= '" + adc("UoCurCdFr") + "'", "")
  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(wb_BaseCur, "1", "0") + " = 1 then 'Y' " + _
            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= OmCmCurCd "
  ' Zubin 213
  
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ***** Manali 3.10.0 - 31/03/12 - Pure Wt For 'X'
  wPurityWt = " /(case when OrRmCtg in ('G', 'P', 'S', 'L') then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst " + _
                          "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and PMCd= OrRmCtg), 1) " + _
                 "When OrRmCtg= 'X' then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst RmP  " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " RmP.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmP.RmCd= PValue and PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmB.RmCtg From RmMst RmB where " + IIF(gs_Partition = ctCurrPrtn, " RmB.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmB.RmCd= RmBaseCd), '')), 1) " + _
                "else 1 end) "
   '****************Geeta**************Emr206*********12/04/04
   'qPtrYN  field included in the query
   ' ****** sachin 2.12 - 24/11/05 - Range Size added to the report query ******'
   
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
      
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " OdOmDt >= '" + Trim(adc("UODTFR")) + "'", "")
      wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " OdOmDt <= '" + Trim(adc("UODTTO")) + "'", "")
    '****** Sachin 3.02.0 - Additional Conditions For New Fields
   
   '*** Rm Req Dynamic
   If UCase(adc.MenuCd) = UCase("RepRmReqDyn") Then
    ' Zubin 213 (ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    '*** Jay 3.2.0 [PicNm]
    Dim ws_BagJoin As String, ws_RngSzJn As String, ws_RngSzCnd As String
    
    ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                 IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " ") + _
                 IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
    
    'mw.137 if Upto RngSz option selected then will get RSz join
    If adc("UoYN5") = "Y" Then
      ws_RngSzJn = "Left Outer Join RmRt ForBalQty On ForBalQty.RrTcTyp='RSZ' and ForBalQty.RrCmCtg= 'C' and " + _
                   "ForBalQty.RrCmCd='" + ctSelfCmCd + "' and ForBalQty.RrCtg=TdRmCtg and ForBalQty.RrSCtg=TdRmSCtg " + _
                   "and ForBalQty.RrSTWGrp= '' and ForBalQty.RrCd= '' and ForBalQty.RrDmCtg= '' and ForBalQty.RrLsCtg= '' " + _
                   "and ForBalQty.RrLabMCd= '' And ForBalQty.RrFrLn<= TdRmSz and ForBalQty.RrToLn>= TdRmSz "
      
      ws_RngSzCnd = " and isnull(ForBalQty.RrRngSz,'') = isnull(RmRt.RrRngSz,'') "
    End If
    If mb_ExactSz1 = True Then
        ws_RngSzCnd = " and (case when PValue3 ='Y' then case when PValue2 ='Y' then OrRmPtr else OrPrdQty/(case when OrPrdQty>0 then OrPrdQty else 1 end) end else OrLn1 end)=" + _
                            "  TdRmSz "
        
    End If
    Dim ws_groupby As String
    Dim ws_fld1 As String
    If mb_OnQty Then
      ws_groupby = "OrCoCd, OrTc, OrYy, OrChr, OrNo, OrOdIdNo, OrSr, OrRmCd, OrRmCtg, OrRmSCtg, OrSr "
      ws_fld1 = ",max(OrLn1) as OrLn1,"
      If mb_ExactSz1 Then
        ws_groupby = " OrCoCd, OrTc, OrYy, OrChr, OrNo, OrOdIdNo, OrSr, OrRmCd, OrRmCtg, OrRmSCtg, (case when PValue3 ='Y' then case when PValue2 ='Y' then OrRmPtr else OrPrdQty/(case when OrPrdQty>0 then OrPrdQty else 1 end) end else OrLn1 end) as OrLn1 "
        ws_fld1 = ",OrLn1 as OrLn1,"
      End If
      If adc("UOYN5") = "Y" Then
        ws_groupby = " OrCoCd, OrTc, OrYy, OrChr, OrNo, OrOdIdNo, OrSr, OrRmCd, OrRmCtg, OrRmSCtg, isnull(RmRt.RrRngSz,'') "
        ws_fld1 = ",(case when PValue3 ='Y' then case when PValue2 ='Y' then OrRmPtr else OrPrdQty/(case when OrPrdQty>0 then OrPrdQty else 1 end) end else OrLn1 end) as OrLn1, "
      End If
    End If
    Dim ws_FldName As String
    ws_FldName = IIF(mb_OnQty, "TdRmQty", "TdRmWt")
    Dim ws_GrpByTemp As String
    moCn.CreateTmpTable "#OrdRmWithFilter", ws_GrpByTemp
 
    ws_GrpByTemp = "Select OrdRm.* into #OrdRmWithFilter " + _
               "From OrdRm Join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey", "") + _
               " Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OrPrtKey", "") + _
               " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + " join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey", "") + _
               " Join Param On PTyp = 'RMSCTG' And PMCd = RmCtg And PSCd = RmSCtg " + ws_Join + _
               " Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= (Case When PValue3='N' Then OrLn1 Else OrRmPtr End) and RrToLn>= (Case When PValue3='N' Then OrLn1 Else OrRmPtr End) " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey='" + ctCurrPrtn + "' ", "") + _
               IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "") + IIF(wCnd <> "", " and ", " ") + " convert(decimal(12, 1), (OdPrdQty- OdFgQty))> 0 " + wCrmCnd
    

    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.10.0 - 31/03/12 - Pure Wt For 'X'
    'mw.137 ws_RngSzJn, ws_RngSzCnd added
    '6.1-98 left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls added
    wSqlStrg = "Select " + grpflds + ", OrRmCd, RmQw, RmQwCst, RmDesc, OrRmCtg, OrRmSCtg, IsNull(RrRngSz,'') as RrRngSz, SCTG.PValue3 as qPtrYN, OrLn1, OrLn2, OrLn3, " + ws_RmPtr + ", " + ws_RmQty + ", " + _
               "(Case when (OrRmCtg= 'D' or OrRmCtg= 'C') then " + ws_RmWt + " else 0.0 end) as qOrCarats, " + _
               "(Case when (OrRmCtg= 'G' or OrRmCtg= 'P' or OrRmCtg= 'S' or OrRmCtg= 'L' or OrRmCtg= 'M' or OrRmCtg= 'X') then " + ws_RmWt + " else 0.0 end) as qOrWt, " + _
               "((Case when (OrRmCtg='G' or OrRmCtg='P' or OrRmCtg='S' or OrRmCtg='L') then RmPurityWt When OrRmCtg='X' Then RmBasePurityWt Else 0.0 End) " + wPurityWt + " )* " + ws_RmWt + " as qOrPureWt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " as OrSalRt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrCstVal/OmCnvFct else OrCstVal*OmCnvFct end) ", "OrCstVal") + " as OrCstRt, " + _
               "OdPrdQty, (OdPrdQty - isnull((Select sum(BQty) from Bag " + ws_BagJoin + " Where " + _
               IIF(gs_Partition = ctCurrPrtn, " BPrtKey=OrPrtKey and ", "") + " BOdIdNo=OrOdIdNo and BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + _
               "and ((BCls= 'Y') or (BLoc= '" + ctFgLoc + "') or (  /*exists(Select 'x' from Txnd where TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdBIdNo=BIdNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=BPrtKey ", "") + _
               "and TdRmCd= OrRmCd) and*/ isnull((select sum(" + ws_FldName + "*( " + _
                                              "Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 Else 0 end)) " + _
               "From Txnd join Param on PTyp= 'RMSCTG' and PMCd= OrRmCtg and PSCd= OrRmSCtg " + ws_RngSzJn + " where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey=BPrtKey and ", "") + " TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd " + ws_RngSzCnd + ") " + _
               ", 0.0)> 0)  )), 0.0)) as qOdBalPrd, OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo, OdDmCd, " + _
               "OdDmSz, OdSfx, OdPicNm, OmCmCd, OmPoNo, DmTcTyp, DmCtg, OmDt, OdDelDt " + _
               "From OrdRm Join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey", "") + _
               " Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OrPrtKey", "") + _
               " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + " join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey", "") + _
               " Join Param SCTG On PTyp = 'RMSCTG' And PMCd = RmCtg And PSCd = RmSCtg " + ws_Join + _
               " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
               " Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) and RrToLn>= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey='" + ctCurrPrtn + "' ", "") + _
               IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "") + IIF(wCnd <> "", " and ", " ") + " convert(decimal(12, 1), (OdPrdQty- OdFgQty))> 0 " + wCrmCnd
    
    If mb_OnQty Then
    wSqlStrg = "Select " + grpflds + ", OrRmCd, RmQw, RmQwCst, RmDesc, OrRmCtg, OrRmSCtg, IsNull(RrRngSz,'') as RrRngSz, SCTG.PValue3 as qPtrYN, OrLn1, OrLn2, OrLn3, " + ws_RmPtr + ", " + ws_RmQty + ", " + _
               "(Case when (OrRmCtg= 'D' or OrRmCtg= 'C') then " + ws_RmWt + " else 0.0 end) as qOrCarats, " + _
               "(Case when (OrRmCtg= 'G' or OrRmCtg= 'P' or OrRmCtg= 'S' or OrRmCtg= 'L' or OrRmCtg= 'M' or OrRmCtg= 'X') then " + ws_RmWt + " else 0.0 end) as qOrWt, " + _
               "((Case when (OrRmCtg='G' or OrRmCtg='P' or OrRmCtg='S' or OrRmCtg='L') then RmPurityWt When OrRmCtg='X' Then RmBasePurityWt Else 0.0 End) " + wPurityWt + " )* " + ws_RmWt + " as qOrPureWt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " as OrSalRt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrCstVal/OmCnvFct else OrCstVal*OmCnvFct end) ", "OrCstVal") + " as OrCstRt, " + _
               "OdPrdQty, (OdPrdQty - isnull((Select sum(BQty) from Bag " + ws_BagJoin + " Where " + _
               IIF(gs_Partition = ctCurrPrtn, " BPrtKey=OrPrtKey and ", "") + " BOdIdNo=OrOdIdNo and BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + _
               "and ((BCls= 'Y') or (BLoc= '" + ctFgLoc + "') or (  /*exists(Select 'x' from Txnd where TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdBIdNo=BIdNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=BPrtKey ", "")
               
    wSqlStrg = wSqlStrg + " " + _
               "and TdRmCd= OrRmCd) and*/ isnull((select sum(" + ws_FldName + "*( " + _
                                              "Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
                                              "Else 0 end)) " + _
               "From Txnd join Param on PTyp= 'RMSCTG' and PMCd= OrRmCtg and PSCd= OrRmSCtg " + ws_RngSzJn + " where " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey=BPrtKey and ", "") + " TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd " + ws_RngSzCnd + ") " + _
               ", 0.0)= OrQty*Bqty ))), 0.0)) as qOdBalPrd, OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo, OdDmCd, " + _
               "OdDmSz, OdSfx, OdPicNm, OmCmCd, OmPoNo, DmTcTyp, DmCtg, OmDt, OdDelDt ," + _
               " (select isnull((select sum(" + ws_FldName + "*( " + _
                                              "Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
                                              "Else 0 end)) " + _
               "From Bag join Txnd on TdBIdno=BidNo join Param on PTyp= 'RMSCTG' and PMCd= OrRmCtg and PSCd= OrRmSCtg " + ws_RngSzJn + " where BOdIdNo=OdIdNo and " + IIF(gs_Partition = ctCurrPrtn, " TdPrtKey=BPrtKey and ", "") + " TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd " + ws_RngSzCnd + ") " + _
               ", 0.0)) as qOrQtyIssued "
               
      wSqlStrg = wSqlStrg + _
               " from OrdRm  " + _
               " Join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey", "") + _
               " Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OrPrtKey", "") + _
               " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + " join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey", "") + _
               " Join Param SCTG On PTyp = 'RMSCTG' And PMCd = RmCtg And PSCd = RmSCtg " + ws_Join + _
               " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
               " Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) and RrToLn>= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey='" + ctCurrPrtn + "' ", "") + _
               IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "") + IIF(wCnd <> "", " and ", " ") + " convert(decimal(12, 1), (OdPrdQty- OdFgQty))> 0 " + wCrmCnd
    End If
                              
    Rep.TxtHead.SetText "Raw Material Requirement (Dynamic)"
    
  End If
  '*** Rm Req Dynamic
  
  '****************Geeta**************Emr206*********12/04/04
   'qPtrYN  field included in the query
   ' ****** sachin 2.12 - 24/11/05 - Range Size added to the report query ******'
   
  '*** Rm Req Static
  If UCase(adc.MenuCd) = UCase("RepRmReqStat") Then
    ' Zubin 213 (ws_Join added)
    '*** Jay 3.2.0 [PicNm]
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali 3.10.0 - 31/03/12 - Pure Wt For 'X'
    '6.1-58 if Omtc is JSO then it gets OdOrdQty else OdPrdQty. existing behaviour- OdPrdQty.
    '6.1-98 left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls  added
    wSqlStrg = "Select " + grpflds + ", " + _
               "OrRmCd, RmQw, RmQwCst, RmDesc, OrRmCtg, OrRmSCtg, IsNull(RrRngSz,'') as RrRngSz, SCTG.PValue3 as qPtrYN," + _
               "OrLn1, OrLn2, OrLn3, " + ws_RmPtr + ", " + ws_RmQty + ", " + _
               "(Case  when (OrRmCtg = 'D' or OrRmCtg = 'C') then " + ws_RmWt + " else 0.0 end) as qOrCarats, " + _
               "(Case  when (OrRmCtg = 'G' or OrRmCtg = 'P' or OrRmCtg = 'S' or OrRmCtg = 'L' or OrRmCtg = 'M' or OrRmCtg = 'X') then " + ws_RmWt + " else 0.0 end) as qOrWt, " + _
               "((Case when (OrRmCtg ='G' or OrRmCtg ='P' or OrRmCtg ='S' or OrRmCtg ='L') then RmPurityWt When OrRmCtg='X' Then RmBasePurityWt Else 0.0 End) " + wPurityWt + ")   * " + ws_RmWt + " as qOrPureWt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + " as OrSalRt, " + _
               IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrCstVal/OmCnvFct else OrCstVal*OmCnvFct end) ", "OrCstVal") + " as OrCstRt, " + _
               IIF(adc("UOOMTCFR") = "JSO", "OdOrdQty as OdPrdQty", "OdPrdQty") + "," + IIF(adc("UOOMTCFR") = "JSO", "OdOrdQty", "OdPrdQty") + " as qOdBalPrd, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo, " + _
               "OdDmCd, OdDmSz, OdSfx, OdPicNm, OmCmCd, OmPoNo, DmTcTyp, DmCtg, OmDt, OdDelDt " + _
               "From OrdRm Join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy " + _
               " and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey", "") + _
               " Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy " + _
               " and OmChr= OrChr and OmNo= OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OrPrtKey", "") + " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' " + _
               " and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + _
               " join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey", "") + _
               " Join Param SCTG On PTyp = 'RMSCTG' And PMCd = RmCtg And PSCd = RmSCtg " + ws_Join + _
               " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
               " Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
               " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) and RrToLn>= (Case When SCTG.PValue3='N' Then OrLn1 Else OrRmPtr End) " + wCnd + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey='" + ctCurrPrtn + "' ", "") + _
               IIF(wCnd <> "", " ", " where 1=1 ") + wCrmCnd + IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " ") + wAddlCnd, "")
    
    ''' Bef 3.10.0
    '''"(Case when (OrRmCtg ='G' or OrRmCtg ='P' or OrRmCtg ='S' or OrRmCtg ='L') then (Select RmPurityWt from RmMst where RmCd=OrRmCd)" + wPurityWt + "  else 0.0 end) * " + ws_RmWt + " as qOrPureWt, " + _

    Rep.TxtHead.SetText "Raw Material Requirement (Static)"
    Rep.hBalPrd1.Suppress = True
    Rep.hBalPrd2.Suppress = True
    Rep.FldDetBalPrd.Suppress = True
    Rep.Fldr1BalPrd.Suppress = True
    Rep.Fldr2BalPrd.Suppress = True
    Rep.Fldr3BalPrd.Suppress = True
    Rep.FldrBalPrd.Suppress = True
  End If
  '*** Rm Req Static
  'Manoj 2.10.0
  With Rep
    If UCase$(adc("UoRmCtgFr")) = "D" Or UCase$(adc("UoRmCtgFr")) = "C" Then
      .PHDPureWt.Suppress = True: .PHDRmWt.Suppress = True
      .PHGPureWt.Suppress = True: .PHGRmWt.Suppress = True
      .DPureWt.Suppress = True: .DRmWt.Suppress = True
      .G3FPureWt.Suppress = True: .G3FRmWt.Suppress = True
      .G2FPureWt.Suppress = True: .G2FRmWt.Suppress = True
      .G1FPureWt.Suppress = True: .G1FRmWt.Suppress = True
      .RFPureWt.Suppress = True: .RFRmWt.Suppress = True
    Else
      .PHDPcRmQty.Suppress = True: .PHDPcCarats.Suppress = True
      .PHGPcRmQty.Suppress = True: .PHGPcCarats.Suppress = True
      .DPcRmQty.Suppress = True: .DPcCarats.Suppress = True
      '****** Sachin 2.14.0 - [02-12-06] ******
      .G1PcRmQty.Suppress = True: .G1FDPcCarats.Suppress = True
      .G2PcRmQty.Suppress = True: .G2FDPcCarats.Suppress = True
      .G3PcRmQty.Suppress = True: .G3FDPcCarats.Suppress = True
      .RFPcRmQty.Suppress = True: .RFDPcCarats.Suppress = True
      '****** Sachin 2.14.0 - [02-12-06] ******
    End If
  End With
  ShowRepSummFld Rep.Fldr1BalPrd, Rep.Fldr2BalPrd, Rep.Fldr3BalPrd, Rep.FldrBalPrd, GRP_REP.GrpOpts, "Rm Code,Rm Ctg,Rm SubCtg,Sub Shp,Parent Cd,Rm Len/Sv,Rm Len 2,Rm Ptr,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt"
  ShowRepSummFld Rep.G1FPrdQty, Rep.G2FPrdQty, Rep.G3FPrdQty, Rep.RFPrdQty, GRP_REP.GrpOpts, "Rm Code,Rm Ctg,Rm SubCtg,Sub Shp,Parent Cd,Rm Len/Sv,Rm Len 2,Rm Ptr,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt"
  'Manoj 2.10.0
  ' **** Zubin 211 **** '
  ' ***** Manali 3.8.0 - Order By added in conctenated String (otherwise speed issue in case of Clustered indexes)
  ws_OrdBy = " Order By OrTc+OrYy+OrChr+Str(OrNo, 10)+Str(OrSr, 10) "
  'ws_OrdBy = " Order By OrTc, OrYy, OrChr, OrNo, OrSr "
  wSqlStrg = wSqlStrg + ws_OrdBy
  ' **** Zubin 211 **** '
  'mw.137 if upto range size option selected then rm range size added for suppress balance qty
  GRP_REP.SetFormula Rep, "r_CndForBalQty", IIF(adc("UoYN5") = "Y", "{rdo.Grp1}+{rdo.Grp2}+{rdo.Grp3}+{rdo.OrTc}+{rdo.OrYy}+{rdo.OrChr}+ ToText({rdo.OrNo}, '######0')+ ToText({rdo.OrSr}, '###0')+{rdo.OrRmCd}+{rdo.RrRngSz}", "{rdo.Grp1}+{rdo.Grp2}+{rdo.Grp3}+{rdo.OrTc}+{rdo.OrYy}+{rdo.OrChr}+ ToText({rdo.OrNo}, '######0')+ ToText({rdo.OrSr}, '###0')+{rdo.OrRmCd}")
   GRP_REP.SetFormula Rep, "wOrQty", IIF(mb_OnQty, "({rdo.orqty}*{rdo.odprdqty})-{rdo.qOrQtyIssued}", "{rdo.OrQty}*{rdo.qOdBalPrd}")
   GRP_REP.SetFormula Rep, "wOrWt", IIF(mb_OnQty, "(({rdo.orqty}*{rdo.odprdqty})-{rdo.qOrQtyIssued})*({rdo.qOrcarats}/iif({rdo.OrQty}>0,{rdo.OrQty},1))", "{rdo.qOrWt}*{rdo.qOdBalPrd}")
   GRP_REP.SetFormula Rep, "wOrCarats", IIF(mb_OnQty, "(({rdo.orqty}*{rdo.odprdqty})-{rdo.qOrQtyIssued})*({rdo.qOrcarats}/iif({rdo.OrQty}>0,{rdo.OrQty},1))", "{rdo.qOrCarats}*{rdo.qOdBalPrd}")
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  If wDetPos = 0 Then
    Rep.PicSec2.Suppress = True
  End If
  
  If adc("UoYN2") = "N" Or wDetPos = 0 Then
    Rep.PICSEC.Suppress = True
    Rep.FldRmdesc.Suppress = True
  Else
    If adc("UoYN4") = "Y" Then
        Rep.picdes.Suppress = True
        Rep.picdesgn.Height = Rep.picdesgn.Height * 2
    Else
        Rep.picdesgn.Suppress = True
    End If
    '        Rep.Picdesgn.YScaling = 2
'        Rep.picdesgn.Height = Rep.picdesgn.Height * 2
'        Rep.picdesgn.Width = Rep.picdesgn.Width * 2
  End If
  
  '****** Sachin 2.14.0 - [Multi-Company] ******
  
  ' Zubin 213
  If wb_BaseCur Then
    Rep.TxtValIn.SetText "Value In " + moCn.GetFldVal("Select HBaseCurCd from Head " + _
                                        "where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + adc("UoCoCdFr") + "'") + ""
  Else
    Rep.TxtValIn.SetText "Value In " + UCase(adc("UoCurCdFr"))
  End If

  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                    "HCd='" + ctSelfCmCd + "'") = "N" Then _
    Rep.hOrRmPtr.Suppress = True: Rep.FldOrRmPtr.Suppress = True
  
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OrRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
 Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'mw.137 check for Range size
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  Case Is = 2     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
    
    If .Value = Checked Then ChkBoxArr(4).Enabled = True Else ChkBoxArr(4).Enabled = False
  Case Is = 3
    If (ChkBoxArr(3).Value = Checked) Then
          mb_OnQty = True
    Else
          mb_OnQty = False
    End If

Case Is = 1
    If (ChkBoxArr(1).Value = Checked) Then
        mb_ExactSz1 = True
    Else
        mb_ExactSz1 = False
    End If
  Case Is = 4     'Tag = "ShowBigPic"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

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
    Set adc.FirNKeyCtl = adc("UoRmCtgFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gb_RmCtgFor = True
    gb_RmZ = False
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "SO"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    'Sachin 3.11.0
    If UCase(adc.MenuCd) <> UCase("RepRmReqDyn") Then
        LblBagTyp.Visible = False
        adc("UoBagTyp").Visible = False
        LblUoBChrFr.Visible = False
        adc("UoBChrFr").Visible = False
        adc("UoBagTyp") = "P": adc("UoBChrFr") = ""
        'mw.137 Range size option visible only for dynamic report
        ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False
        ChkBoxArr(1).Enabled = False: ChkBoxArr(1).Visible = False
        ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False
    End If
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  Select Case UCase(IdName)
  ' urmi 2.06 Pwd for Val
  Case Is = UCase("UoYn")
    '*** Jay 2.14 *** (SEO Change)
      If pv_NewValue <> "" Then
         If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") Then _
              Cancel = True: ErrMsg = "Enter SEO Password to View Values ": Exit Sub
      End If
' Bef 2.14
'    If pv_NewValue <> "" And Trim(TxtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'      Cancel = True: DispMsg "Enter SEO Password to View Values", etError
'    End If
' Bef 2.14
    '*** Jay 2.14 *** (SEO Change)
    
    '*** Sachin 2.11.0
'    If pv_NewValue <> "" Then
'      If adc("UoYn3") = "Y" Then
'        If pv_NewValue <> "Y" Then Cancel = True: ErrMsg = "If Prd RM and Qty is 'Y' then Cost Value should be Selected"
'      Else
'        If pv_NewValue <> "N" Then Cancel = True: ErrMsg = "If Prd RM and Qty is 'N' then Sale Value should be Selected"
'      End If
'    End If
    '*** Sachin 2.11.0
        
  Case Is = UCase("UoYn1")
    If pv_NewValue <> "" And Not moCn.RecSeek("Select 1 from Param where " + _
                           "Ptyp = 'YN' and PmCd = '" + pv_NewValue + "'") Then
      Cancel = True: ErrMsg = "Value Not Defined. (Enter 'Y'es/'N'o)"
    End If
  
  Case Is = UCase("UoYn3")
    If pv_NewValue <> "" And Not moCn.RecSeek("Select 1 from Param where " + _
                           "Ptyp = 'YN' and PmCd = '" + pv_NewValue + "'") Then
      Cancel = True: ErrMsg = "Value Not Defined. (Enter 'Y'es/'N'o)"
    End If
    
  Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
  
  '578-28 validating order no. selection should be a number
  Case Is = UCase("UoOmNoSel")
  Dim wOmNoSel() As String, i As Integer
    wOmNoSel = Split(pv_NewValue, ",")
    If IsArray(wOmNoSel) Then
      For i = 0 To UBound(wOmNoSel)
        Cancel = Not IsNumeric(wOmNoSel(i))
        If Cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is Not Numeric. Enter Order Numbers like 1,2,3 ": Exit Sub
      Next i
    End If
          
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True   '*** (Jen 2.14 Next Patch 1)
  
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      Cancel = IIF(adc("UoYN3") = "N", True, False)
      ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used"
      
    Case Is = UCase("UoYn4")
      Cancel = IIF(adc("UoYN2") = "N", True, False)
      ErrMsg = "Can be Used Only When Show Picture is 'Y'"
      
    Case Is = UCase("UoBChrFr")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If

    Case Is = UCase("UoPMCdFr"), UCase("UoPMCdTo"), UCase("UoPMCdSel")
    gs_Tbl = "vParam"
    gs_PTyp = "PRDSEQ"
    
    If adc.Mode = xNorm Then
      If IdName = UCase("UoPMCdSel") Then HlpList.MultiSelect = True
      HlpList.vPMCd adc("UoCoCdFr"), gs_PTyp: Exit Sub
    End If
    
    
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      If adc("UoYN3") = "N" Then adc("UoYn1") = "N": OptPtr(1).Value = True
    
    '*** Sachin 2.11.0
    Case Is = UCase("UoYn")
       If adc("UoYn") <> "" Then adc("UoYn") = adc("UoYn3")
    '*** Sachin 2.11.0
    
    '***** Sachin 3.11.0
    Case Is = UCase("UOBCHRFR")
      adc("UOBCHRFR") = ""

  End Select
End Sub

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptPtr_Click(Index As Integer)
  Select Case Index
  Case Is = 0
      If OptRmPrd(1).Value = True Then
        OptPtr(0).Value = False: OptPtr(1).Value = True: adc("UoYN1") = "N"
        DispMsg "Calc Ptr Can be Selected Only When Prd Wt & Prd Qty is Selected", etError
      Else
        adc("UoYN1") = "Y"
      End If
  Case Is = 1
      adc("UoYN1") = "N"
  End Select
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptPtr_GotFocus(Index As Integer)
  DispMsg FraOptPtr.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
  Select Case Index
  Case Is = 0
      adc("UoYN3") = "Y"
      If OptRmVal(2).Value = False Then OptRmVal(0).Value = True
  Case Is = 1
      adc("UoYN3") = "N"
      OptPtr(1).Value = True
      If OptRmVal(2).Value = False Then OptRmVal(1).Value = True
  End Select
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptRmVal_Click(Index As Integer)
  Select Case Index
  Case Is = 0
     If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") Then
        OptRmVal(0).Value = False: OptRmVal(2).Value = True: adc("UoYN3") = ""
        DispMsg "Enter SEO Password to View Values", etError
     Else
        adc("UoYN") = "Y"
     End If
  Case Is = 1
          If Trim(TxtPwd.text) <> moCn.GetFldVal("select IsNull ((select PValue from Param Join UsrAccess on UaUsrCd= PMCd " + _
              "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "'  and UaMnuCd= '" + adc.MenuCd + "' and " + _
              "UaSeoYn= 'Y' and PValue<> '' and PValue= '" + Trim(TxtPwd.text) + "' ), " + _
              "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' ))") Then
        OptRmVal(1).Value = False: OptRmVal(2).Value = True: adc("UoYN3") = ""
        DispMsg "Enter SEO Password to View Values", etError
     Else
        adc("UoYN") = "N"
     End If
  Case Is = 2
     adc("UoYN") = ""
  End Select
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
Private Sub OptRmVal_GotFocus(Index As Integer)
  DispMsg FraOptRmVal.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added

'urmi 2.06
Private Sub TxtPwd_Validate(Cancel As Boolean)
'*** Jay 2.14 *** (SEO Change)
  Dim ws_VldPwd As String
  ws_VldPwd = moCn.GetFldVal("select isnull ((select PValue from Param Join UsrAccess on UaUsrCd = PMCd " + _
                   "where PTyp= 'USR' and PMCd= '" + gs_UsrCd + "' and UaMnuCd = '" + adc.MenuCd + "' and " + _
                   "UaSeoYn= 'Y' and PValue= '" + Trim(TxtPwd.text) + "'), " + _
                   "(Select PValue1 from Param where Ptyp= 'SEOPWD' and PmCd= 'SEOPWD' )) ")
    
  If Trim(TxtPwd.text) = "" And ws_VldPwd <> "" Then
    adc("UoYN") = "": OptRmVal(2).Value = True
    Exit Sub  ' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
  Else

  End If
  If Trim(TxtPwd.text) <> "" And Trim(TxtPwd.text) <> ws_VldPwd Then
      Cancel = True:  DispMsg "Invalid Password / User Does Not Have Seo Rights ", etError
  End If
' Bef 2.14
' If Trim(TxtPwd.text) = "" Then adc("UoYn") = ""
'
'    If Trim(TxtPwd.text) = "" Or Trim(TxtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
'                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
'
'      Exit Sub
'    Else
'      Cancel = True
'      DispMsg "Invalid Password", etError
'    End If
' Bef 2.14
'*** Jay 2.14 *** (SEO Change)
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
 adc("UOYN") = "": OptRmVal(2).Value = True
 
' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
If adc("UOYN3") = "Y" Then OptRmPrd(0).Value = True
If adc("UOYN3") = "N" Then OptRmPrd(1).Value = True
If adc("UOYN") = "Y" Then OptRmVal(0).Value = True
If adc("UOYN") = "N" Then OptRmVal(1).Value = True
If adc("UOYN") = "" Then OptRmVal(2).Value = True
If adc("UOYN1") = "Y" Then OptPtr(0).Value = True
If adc("UOYN1") = "N" Then OptPtr(1).Value = True
  ' ****** Manali 3.5.0 - 04/11/08 - Radio Buttons Added
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  
  '***** Sachin 3.11.0 - Bag Type in Requirement Report
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChrFr").Visible = False
    adc("UoBagTyp") = "P": adc("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing
  'mw.137 default value for range sz scope is No
  adc("UoYN5") = "N"
  ChkBoxArr(0).Value = Unchecked

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
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

Private Sub xxx()
'    Dim wStr As String, wArr() As String, i As Integer, OptPos As Integer, wSupGrand As Boolean
'    wStr = GRP_REP.GrpOpts
'    If InStr(1, "," + wStr + ",", "," + "Kt" + ",") = 0 Then
'        t3_netwt.Suppress = True: t2_netwt.Suppress = True: t1_netwt.Suppress = True: g_netwt.Suppress = True
'    Else
'        wArr = Split(wStr)
'        For i = 0 To 2
'           If wArr(i) = "Kt" Then
'              OptPos = i
'              Exit For
'            End If
'        Next i
'        wSupGrand = False
'        For i = 0 To 2
'          If i < OptPos Then
'            Dim xxx As CRAXDRT.Report
'            xxx.Sections("G" + CStr(i) + "F").ReportObjects("T" + CStr(i) + "NetWt").Suppress = True
'            wSupGrand = True
'          End If
'        Next i
'        g_netwt.Suppress = True
'    End If
End Sub





