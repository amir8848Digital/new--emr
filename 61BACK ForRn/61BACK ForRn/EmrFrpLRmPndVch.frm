VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpLRmPndVch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Voucher (Loose Rm)"
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
      Left            =   11070
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   67
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
      Left            =   5175
      TabIndex        =   75
      Top             =   9660
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   76
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
         TabIndex        =   77
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
      Left            =   6390
      TabIndex        =   74
      Top             =   9675
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   71
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
      TabIndex        =   68
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
      Height          =   9735
      Left            =   0
      TabIndex        =   70
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   69
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpLRmPndVch.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpLRmPndVch.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   72
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraRmScp 
               BorderStyle     =   0  'None
               Height          =   2055
               Left            =   0
               TabIndex        =   101
               Top             =   5040
               Width           =   15105
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   82
                  Left            =   4650
                  TabIndex        =   53
                  ToolTipText     =   "Enter Raw Material Code To Range"
                  Top             =   570
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataField       =   "UORMCDTO"
                  IdName          =   "UORMCDTO"
                  CmpStr          =   "RmCd <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   83
                  Left            =   7710
                  TabIndex        =   54
                  ToolTipText     =   "Enter Raw Material Code Selection"
                  Top             =   570
                  Width           =   7290
                  _ExtentX        =   12859
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCDSEL"
                  IdName          =   "UORMCDSEL"
                  CmpStr          =   "RmCd In"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   84
                  Left            =   4650
                  TabIndex        =   50
                  ToolTipText     =   "Enter Raw Material Sub Category To Range"
                  Top             =   285
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
                  Index           =   85
                  Left            =   7710
                  TabIndex        =   51
                  ToolTipText     =   "Enter Raw Material Sub Category Selection"
                  Top             =   285
                  Width           =   7290
                  _ExtentX        =   12859
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMSCTGSEL"
                  IdName          =   "UORMSCTGSEL"
                  CmpStr          =   "RmSCtg In"
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   86
                  Left            =   1605
                  TabIndex        =   46
                  ToolTipText     =   "Enter From Raw Material Category"
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UORMCTGFR"
                  IdName          =   "UORMCTGFR"
                  CmpStr          =   "RmCtg >= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   87
                  Left            =   1605
                  TabIndex        =   49
                  ToolTipText     =   "Enter Raw Material Sub Category From Range"
                  Top             =   285
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
                  Index           =   88
                  Left            =   1605
                  TabIndex        =   52
                  ToolTipText     =   "Enter Raw Material Code From Range"
                  Top             =   570
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataField       =   "UORMCDFR"
                  IdName          =   "UORMCDFR"
                  CmpStr          =   "RmCd >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   89
                  Left            =   4650
                  TabIndex        =   56
                  ToolTipText     =   "Enter To Sale Rate"
                  Top             =   855
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOSALRTTO"
                  IdName          =   "UOSALRTTO"
                  CmpStr          =   "OdJLVchRt <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   90
                  Left            =   1605
                  TabIndex        =   55
                  ToolTipText     =   "Enter From Sale Rate"
                  Top             =   855
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  MaxLength       =   0
                  DataType        =   2
                  DataField       =   "UOSALRTFR"
                  IdName          =   "UOSALRTFR"
                  CmpStr          =   "OdJLVchRt >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   91
                  Left            =   4650
                  TabIndex        =   47
                  ToolTipText     =   "Enter To Raw Material Category"
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UORMCTGTO"
                  IdName          =   "UORMCTGTO"
                  CmpStr          =   "RmCtg <= "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   92
                  Left            =   7710
                  TabIndex        =   48
                  ToolTipText     =   "Enter Raw Material Category Selection"
                  Top             =   0
                  Width           =   7290
                  _ExtentX        =   12859
                  _ExtentY        =   503
                  MaxLength       =   255
                  DataField       =   "UORMCTGSEL"
                  IdName          =   "UORMCTGSEL"
                  CmpStr          =   "RmCtg In "
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   93
                  Left            =   4650
                  TabIndex        =   58
                  ToolTipText     =   "Enter To Raw Material Len/Sv"
                  Top             =   1140
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN1TO"
                  IdName          =   "UOLN1TO"
                  CmpStr          =   "OdJLLn1 <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   94
                  Left            =   1605
                  TabIndex        =   57
                  ToolTipText     =   "Enter From Raw Material Len/Sv"
                  Top             =   1140
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN1FR"
                  IdName          =   "UOLN1FR"
                  CmpStr          =   "OdJLLn1 >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   2
                  Left            =   4650
                  TabIndex        =   60
                  ToolTipText     =   "Enter To Raw Material Len/Sv"
                  Top             =   1425
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN2TO"
                  IdName          =   "UOLN2TO"
                  CmpStr          =   "OdJLLn2 <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   15
                  Left            =   1605
                  TabIndex        =   59
                  ToolTipText     =   "Enter From Raw Material Len/Sv"
                  Top             =   1425
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN2FR"
                  IdName          =   "UOLN2FR"
                  CmpStr          =   "OdJLLn2 >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   16
                  Left            =   4650
                  TabIndex        =   62
                  ToolTipText     =   "Enter To Raw Material Len/Sv"
                  Top             =   1710
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN3TO"
                  IdName          =   "UOLN3TO"
                  CmpStr          =   "OdJLLn3 <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   33
                  Left            =   1605
                  TabIndex        =   61
                  ToolTipText     =   "Enter From Raw Material Len/Sv"
                  Top             =   1710
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  Alignment       =   2
                  Mask            =   "#0.000"
                  MaxLength       =   6
                  DataType        =   2
                  DataField       =   "UOLN3FR"
                  IdName          =   "UOLN3FR"
                  CmpStr          =   "OdJLLn3 >="
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
                  Index           =   7
                  Left            =   90
                  TabIndex        =   111
                  Top             =   1425
                  Width           =   1425
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
                  Index           =   0
                  Left            =   90
                  TabIndex        =   110
                  Top             =   1710
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
                  Index           =   32
                  Left            =   90
                  TabIndex        =   106
                  Top             =   1140
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
                  Index           =   22
                  Left            =   90
                  TabIndex        =   105
                  Top             =   570
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
                  Index           =   6
                  Left            =   90
                  TabIndex        =   104
                  Top             =   285
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
                  Index           =   5
                  Left            =   90
                  TabIndex        =   103
                  ToolTipText     =   "Location"
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.Label ALBL 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rm Rate"
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
                  Left            =   90
                  TabIndex        =   102
                  Top             =   855
                  Width           =   1425
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Only Pending      "
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
               Left            =   4650
               TabIndex        =   43
               Tag             =   "ShowPnd"
               ToolTipText     =   "Check To Show Pending Records"
               Top             =   4455
               Width           =   1890
            End
            Begin VB.Frame FraOptDelDt 
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   6420
               TabIndex        =   99
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date/ Both Dates"
               Top             =   7320
               Width           =   2235
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2265
               TabIndex        =   64
               ToolTipText     =   "Show Box Detail ? (Yes / No)"
               Top             =   7110
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
               Index           =   9
               Left            =   6825
               TabIndex        =   45
               ToolTipText     =   "Show Pending Orders Or All Orders? (Yes / No)"
               Top             =   4440
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
               Index           =   11
               Left            =   1605
               TabIndex        =   42
               ToolTipText     =   "Enter Pending As Of Date"
               Top             =   4455
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
               CmpStr          =   "BOm.OmDt <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   5415
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Year"
               Top             =   960
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "B.OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   2385
               TabIndex        =   5
               ToolTipText     =   "Enter From Order Year"
               Top             =   960
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "B.OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   1605
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Tc"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               ReCalcOn        =   "UOOMTC"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "B.OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2805
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Character"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "B.OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4650
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Tc"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "B.OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   5835
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Character"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "B.OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5535
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1245
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "B.OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   2505
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1245
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "B.OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   1605
               TabIndex        =   10
               ToolTipText     =   "Enter From Order Number"
               Top             =   1245
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "B.OdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4650
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Number"
               Top             =   1245
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "B.OdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4650
               TabIndex        =   15
               ToolTipText     =   "Enter To Order Date"
               Top             =   1530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "BOm.OmDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1605
               TabIndex        =   14
               ToolTipText     =   "Enter From Order Date"
               Top             =   1530
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "BOm.OmDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4650
               TabIndex        =   17
               ToolTipText     =   "Enter To Customer"
               Top             =   1815
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
               CmpStr          =   "BOm.OmCmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7710
               TabIndex        =   18
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1815
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "BOm.OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1605
               TabIndex        =   16
               ToolTipText     =   "Enter From Customer"
               Top             =   1815
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
               CmpStr          =   "BOm.OmCmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4650
               TabIndex        =   26
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2700
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
               CmpStr          =   "BOm.OmPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   7710
               TabIndex        =   27
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2700
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
               CmpStr          =   "BOm.OmPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   1605
               TabIndex        =   25
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2700
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
               CmpStr          =   "BOm.OmPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   4650
               TabIndex        =   37
               ToolTipText     =   "Enter To Export Delivery Date"
               Top             =   3600
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "B.OdExpDelDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1605
               TabIndex        =   36
               ToolTipText     =   "Enter From Export Delivery Date"
               Top             =   3600
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "B.OdExpDelDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   4650
               TabIndex        =   39
               ToolTipText     =   "Enter To Balance Quantity"
               Top             =   3885
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDTO"
               IdName          =   "UOBALPRDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1605
               TabIndex        =   38
               ToolTipText     =   "Enter From Balance Quantity"
               Top             =   3885
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOBALPRDFR"
               IdName          =   "UOBALPRDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   4650
               TabIndex        =   41
               ToolTipText     =   "Enter To Balance Value"
               Top             =   4170
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTTO"
               IdName          =   "UOCSTRTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   1605
               TabIndex        =   40
               ToolTipText     =   "Enter From Balance Value"
               Top             =   4170
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UOCSTRTFR"
               IdName          =   "UOCSTRTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   1605
               TabIndex        =   44
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   4740
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
               Index           =   75
               Left            =   5930
               TabIndex        =   31
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   2985
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "B.OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   7710
               TabIndex        =   32
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   2985
               Visible         =   0   'False
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
               CmpStr          =   "B.OdPoNo In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   77
               Left            =   2900
               TabIndex        =   29
               ToolTipText     =   "Enter From Sub Purchase Order "
               Top             =   2985
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "B.OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   80
               Left            =   1605
               TabIndex        =   3
               ToolTipText     =   "Enter Base Order Tc"
               Top             =   675
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
               Index           =   98
               Left            =   7710
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   390
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "B.OdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   99
               Left            =   7710
               TabIndex        =   21
               ToolTipText     =   "Enter Customer Group Selection"
               Top             =   2115
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMGRPSEL"
               IdName          =   "UOCMGRPSEL"
               CmpStr          =   "CmGrp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   100
               Left            =   4650
               TabIndex        =   20
               ToolTipText     =   "Enter To Customer Group"
               Top             =   2115
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPTO"
               IdName          =   "UOCMGRPTO"
               CmpStr          =   "CmGrp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   101
               Left            =   1605
               TabIndex        =   19
               ToolTipText     =   "Enter From Customer Group"
               Top             =   2115
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMGRPFR"
               IdName          =   "UOCMGRPFR"
               CmpStr          =   "CmGrp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   102
               Left            =   7710
               TabIndex        =   24
               ToolTipText     =   "Enter Customer Type Selection"
               Top             =   2415
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMTYPSEL"
               IdName          =   "UOCMTYPSEL"
               CmpStr          =   "CmTyp In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   103
               Left            =   4650
               TabIndex        =   23
               ToolTipText     =   "Enter To Customer Type"
               Top             =   2415
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPTO"
               IdName          =   "UOCMTYPTO"
               CmpStr          =   "CmTyp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   104
               Left            =   1605
               TabIndex        =   22
               ToolTipText     =   "Enter From Customer Type"
               Top             =   2415
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMTYPFR"
               IdName          =   "UOCMTYPFR"
               CmpStr          =   "CmTyp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   1605
               TabIndex        =   33
               ToolTipText     =   "Enter From Lot No."
               Top             =   3270
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UOLOTNOFR"
               IdName          =   "UOLOTNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4650
               TabIndex        =   34
               ToolTipText     =   "Enter To Lot No."
               Top             =   3270
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   16
               DataType        =   4
               DataField       =   "UOLOTNOTO"
               IdName          =   "UOLOTNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   7710
               TabIndex        =   35
               ToolTipText     =   "Enter Lot No. Selection"
               Top             =   3270
               Width           =   7290
               _ExtentX        =   12859
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOLOTNOSEL"
               IdName          =   "UOLOTNOSEL"
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Box Det (Y/N)"
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
               Left            =   60
               TabIndex        =   63
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show box Details"
               Top             =   7200
               Width           =   1770
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1605
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   390
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "B.OdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   97
               Left            =   4650
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   390
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "B.OdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   4440
               TabIndex        =   112
               ToolTipText     =   "Show Export/Production/Both Delivery Dates? (Enter ""Y"" for Export , ""N"" for Production and ' ' for Both)"
               Top             =   7200
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
               Data            =   "N"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   65
               Left            =   1605
               TabIndex        =   28
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   3000
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
               Index           =   66
               Left            =   4650
               TabIndex        =   30
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   3000
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
               CmpStr          =   "OdSubCust <="
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
               Index           =   35
               Left            =   5860
               TabIndex        =   114
               Top             =   3000
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
               Index           =   36
               Left            =   2820
               TabIndex        =   113
               Top             =   3000
               Width           =   105
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
               TabIndex        =   109
               Top             =   3270
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust/Supp Typ"
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
               Left            =   60
               TabIndex        =   108
               Top             =   2410
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust/Supp Grp"
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
               Left            =   60
               TabIndex        =   107
               Top             =   2120
               Width           =   1425
            End
            Begin VB.Label LblBaseTc 
               BackStyle       =   0  'Transparent
               Caption         =   "Base Tc"
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
               TabIndex        =   100
               Top             =   675
               Width           =   1605
            End
            Begin VB.Label ALBL 
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
               Index           =   28
               Left            =   60
               TabIndex        =   98
               Top             =   2985
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
               Left            =   60
               TabIndex        =   97
               Top             =   4740
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
               Index           =   25
               Left            =   60
               TabIndex        =   96
               Top             =   390
               Width           =   1335
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Val"
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
               Left            =   60
               TabIndex        =   95
               Top             =   4170
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Wt"
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
               Left            =   60
               TabIndex        =   94
               Top             =   3885
               Width           =   1425
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
               Index           =   20
               Left            =   60
               TabIndex        =   93
               Top             =   3600
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
               TabIndex        =   92
               Top             =   2700
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Cust / Supp"
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
               TabIndex        =   91
               Top             =   1815
               Width           =   1425
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
               Left            =   60
               TabIndex        =   90
               Top             =   1530
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
               Left            =   2445
               TabIndex        =   89
               Top             =   1245
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
               Left            =   5475
               TabIndex        =   88
               Top             =   1245
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
               Left            =   2325
               TabIndex        =   87
               Top             =   960
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
               Left            =   5355
               TabIndex        =   86
               Top             =   960
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
               Left            =   2745
               TabIndex        =   85
               Top             =   960
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
               Left            =   5775
               TabIndex        =   84
               Top             =   960
               Width           =   105
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
               Index           =   11
               Left            =   60
               TabIndex        =   83
               Top             =   1245
               Width           =   1485
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
               Left            =   60
               TabIndex        =   82
               Top             =   960
               Width           =   1785
            End
            Begin VB.Label LblAsOfDt 
               BackStyle       =   0  'Transparent
               Caption         =   "As Of Date"
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
               Left            =   60
               TabIndex        =   81
               Top             =   4455
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sel"
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
               Left            =   7710
               TabIndex        =   80
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
               Left            =   1605
               TabIndex        =   79
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
               Left            =   4650
               TabIndex        =   78
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
            TabIndex        =   73
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   65
               Top             =   360
               Width           =   9975
               _ExtentX        =   17595
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpLRmPndVch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Report --------'
' This report is used to give a detail listing of all the pending orders in the system as
'   on a particular date. The report can also be used
' The user has the option of seeing/ hiding the corresponding matching invoice records.
'--------------------

Option Explicit
Dim Rep As New EmrRepLRmPndVch
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06
Dim ms_PwdFlag As Boolean
Dim ws_MultiCoMod As Boolean


Private Sub SetGroupSort()
  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "B.OdCoCd", "B.OdCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=B.OdCoCd) "
  
  GRP_REP.Add "Voucher No", "B.OdTc+B.OdYy+B.OdChr+str(B.OdNo)", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))", "", "", "Order Cust,Customer PO,Customer Grp,Customer Typ", "", "", "B.OdTc+'/'+B.OdYy+'/'+B.OdChr+'/'+LTrim(Str(B.OdNo))"
  '6.1
  GRP_REP.Add "Voucher Date", "convert(varchar(10), BOm.OmDt, 112)", "convert(varchar(8),BOm.OmDt, " + CStr(DtFmtSlahVal()) + ")", "OmDt", "hOmDt", "Quarter,Month,Week", "", "", ""
'  GRP_REP.Add "Week", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' - '+DateName(WW,BOm.OmDt)", "", "", "", "Quarter,Month", "", "", ""
'  GRP_REP.Add "Month", "Str(DatePart(yyyy,BOm.OmDt))+Str(DatePart(mm,BOm.OmDt))", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' '+DateName(MM,BOm.OmDt)", "", "", "Quarter", "", "", ""
'  GRP_REP.Add "Quarter", "LTrim(Str(DatePart(yyyy,BOm.OmDt)))+' - '+DateName(QQ,BOm.OmDt)", "", "", "", "", "", "", ""
  GRP_REP.Add "Cust / Supp", "BOm.OmCmCd", "BOm.OmCmCd", "OmCmCd", "hOmCmCd", "Customer Grp,Customer Typ", "", "", "(Select ( case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(gs_UsrCd) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg=BOm.OmCmCtg and CmCd= BOm.OmCmCd) "
'  GRP_REP.Add "Customer Grp", "CmGrp", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMGRP' and PmCd = CmGrp)"
'  GRP_REP.Add "Customer Typ", "CmTyp", "", "", "", "", "", "", "(Select PDesc from Param where Ptyp = 'CMTYP' and PmCd = CmTyp)"
 
'    GRP_REP.Add "Prd Del Date", "convert(varchar(10), B.OdDelDt, 112)", "convert(varchar(8), B.OdDelDt, 3)", "", "", "", "", "", ""
'6.1
    GRP_REP.Add "Exp Del Date", "convert(varchar(10), B.OdExpDelDt, 112)", "convert(varchar(8), B.OdExpDelDt, " + CStr(DtFmtSlahVal()) + ")", "", "", "", "", "", ""
 '   GRP_REP.Add "Customer PO", "BOm.OmPoNo", "", "OmPoNo", "HOmPoNo", "Order Cust", "", "", "convert(varchar(8), BOm.OmPoDt, 3)"
 
    GRP_REP.Add "Rm Ctg", "RmCtg", "", "RmCtg", "hRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = RmCtg) "
    GRP_REP.Add "Rm SubCtg", "RmCtg+RmSCtg", "'('+ RTrim(RmCtg)+ ') '+ RmSCtg", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
    GRP_REP.Add "Rm Code", "RmCtg+RmSCtg+B.OdJLRmCd", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ B.OdJLRmCd", "RmCtg,RmSCtg,OdJLRmCd", "hRmCtg,hRmSCtg,hRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "
    GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "RmCtg,RmSCtg", "hRmCtg,hRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  Dim wsPtrYN As String
  wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
  
  'GRP_REP.Add "Rm Sv/Len", "case when " + wsPtrYN + "='Y' then  str(B.OdJLLn1,5,2)+str(B.OdJLLn2,5,2)+str(B.OdJLLn3,5,2) else IsNull(RisRmSzDesc, str(B.OdJLLn1,5,2)+str(B.OdJLLn2,5,2)+str(B.OdJLLn3,5,2)) end ", _
   "", "B.OdJLLn1,B.OdJLLn2,B.OdJLLn3", "hOrLn1,hOrLn2,hOrLn3", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
  
  GRP_REP.Add "Rm Sv/Len", Array("case when " + wsPtrYN + "='Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else RisRmSzDesc end) end ", _
      "case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + "  else RisRmSzDesc end) end ", _
      "case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + "  else RisRmSzDesc end) end "), _
           "", "wRmSz", "hOrLn1", "Rm Ptr,(RmCd) Sv/Len,Range Size", "", "", ""
                      
'"case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + "  else RisRmSzDesc end) end ", _
      "case when " + wsPtrYN + "= 'Y' then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + " else (case when RisRmSzDesc is null then " + GetRmSzDBFldNm("OdJLLn1", "OdJLLn2", "OdJLLn3") + "  else RisRmSzDesc end) end "), _


If (moCn.RecSeek("select 'x' from Head where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + gs_CoCd + "' and HLotNoYN= 'Y' ") And ws_MultiCoMod = False) Or (ws_MultiCoMod = True) Then
    GRP_REP.Add "RM Lot No.", "B.OdJLLotno", "B.OdJLLotNo", "OdJLLotNo", "hLotNo", "", "", "", "B.OdJLLotno"
End If
    
End Sub
Private Sub ADC_NonKeyValid(pr_Cancel As Boolean, pr_ErrMsg As String)
  ' ***** Manali Trading Module
  '###########################################  Manoj  ###########################################
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
'*** Jay 3.2.0 [PicNm]
' ***** Manali 3.5.0 - 11/11/08 - [OdSalRem, qDetPos, qDispSalRemYN] added
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(45) as DGrp1, space(45) as DGrp2, space(45) as DGrp3,
'  space(45) as G1Desc, space(45) as G2Desc, space(45) as G3Desc,
'  OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd,
'  OdSfx, OdDmSz, OdPicNm, OdOrdQty, OdSalPrc, OdKt, OdPrtCd , DmSalCtg,
'  OdExpDelDt, OdJLExpWt, OdDelDt as qDelDt,IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InDt,
'  InExpNo , InCmCd, IdQty, idiFOB, OdSalRem
'  From OrdDsg, OrdMst, DsgMst, InvDsg, InvHd where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wMtchStr As String, wExpWt As String, wAllPndCnd As String, wRecSelFormula As String
  Dim wOrdDt As String
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
  ' ***** Manali Trading Module
  Dim wBaseRefIdKey As String
  
   
  Dim ws_ExpJoin As String, wBalWt As String, wOrdWt As String, wOrdAmdJoin As String, wOdWtCnd As String    ' Manali 3.7.1
  Dim wTcDesc As String
  
  ' Zubin 213
  Dim wb_InBaseCur As Boolean
  Dim ws_Join As String
  wb_InBaseCur = IIF(adc("UoCurCdFr") = "", True, False)
  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(wb_InBaseCur, "1", "0") + " = 1 then 'Y' " + _
            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= BOm.OmCmCurCd "
   
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst

'  grpflds = "BOm.OmCmCurCd "
    
  ' **** Crm 2.12 - 23-02-06 **** '
  Dim wCrmCnd As String
  wCrmCnd = " And BOm.OmSubmitYN <> 'N' "
  ' **** Crm 2.12 - 23-02-06 **** '
    
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
  End If
 
  wBaseRefIdKey = "OdBaseIdKey"
  
  wExpWt = "": wAllPndCnd = "": wRecSelFormula = ""
    
'    wExpWt = " B.OdJLExpWt as OdJLExpWt "
    wBalWt = " B.OdJLWt - B.OdJLExpWt "
    ' **** Manali 3.7.1
    If adc("UoYN") = "Y" Then wAllPndCnd = " and (B.OdCls='N') and B.OdJLWt- B.OdJLExpWt > 0 "

    adc("UoBalPrdFr").CmpStr = "B.OdJLWt - B.OdJLExpWt >= "
    adc("UoBalPrdTo").CmpStr = "B.OdJLWt - B.OdJLExpWt <= "
    ' Zubin 213 (Curr cnd added)
    adc("UoCstRtFr").CmpStr = "(B.OdJLWt - B.OdJLExpWt)* " + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchRt/BOm.OmCnvFct) else (B.OdJLVchRt*BOm.OmCnvFct) end) ", "B.OdJLVchRt") + " >= "
    adc("UoCstRtTo").CmpStr = "(B.OdJLWt - B.OdJLExpWt)* " + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchRt/BOm.OmCnvFct) else (B.OdJLVchRt*BOm.OmCnvFct) end) ", "B.OdJLVchRt") + " <= "
    
    wOdWtCnd = " and B.OdJLWt >0 "
        
      If adc("UoYN") = "Y" Then
        wAllPndCnd = " and Round(" + wBalWt + ", 3)> 0 "
      End If
   
    adc("UoBalPrdFr").CmpStr = "": adc("UoBalPrdTo").CmpStr = ""
    adc("UoCstRtFr").CmpStr = "": adc("UoCstRtTo").CmpStr = ""
    
    wrepcnd = adc.RepCond
    wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
    If adc("UoBalPrdFr") > 0 Then wRecSelFormula = wRecSelFormula + "{@wBalWt}>= " + CStr(adc("UoBalPrdFr")) + " "
    If adc("UoBalPrdTo") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalWt}<= " + CStr(adc("UoBalPrdTo")) + " "
    If adc("UoCstRtFr") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalVal}>= " + CStr(adc("UoCstRtFr")) + " "
    If adc("UoCstRtTo") > 0 Then wRecSelFormula = wRecSelFormula + IIF(wRecSelFormula = "", "", " and ") + "{@wBalVal}<= " + CStr(adc("UoCstRtTo")) + " "
    If wRecSelFormula <> "" Then Rep.RecordSelectionFormula = wRecSelFormula
  
  
  ' Zubin 213
  wCnd = wCnd + IIF(wb_InBaseCur = False, " and CmCurCd= '" + adc("UoCurCdFr") + "'", "")
  
 ' wExpWtAsOfDt = " IsNull((Select sum(IdQty) from InvDsg Id join InvHd Ih on " + _
                IIF(gs_Partition = ctCurrPrtn, " Ih.InPrtKey=Id.IdPrtKey and ", "") + " Ih.InIdNo=Id.IdInIdNo and Ih.InCoCd= Id.IdCoCd and Ih.InTc= Id.IdTc and Ih.InYy= Id.IdYy " + _
                " and Ih.InChr= Id.IdChr and Ih.InNo= Id.IdNo " + _
                " where Id.IdOdIdNo=OdIdNo and  Id.IdCoCd= OdCoCd and Id.IdExpOdTc= OdTc and Id.IdExpOdYy= OdYy " + _
                " and Id.IdExpOdChr= OdChr and Id.IdExpOdNo= OdNo and Id.IdExpOdSr= OdSr  " + IIF(gs_Partition = ctCurrPrtn, " and Id.IdPrtKey=OdPrtKey ", "") + _
               IIF(adc("UoAsOfDt") = "01/01/80", "", " and  Ih.InDt<= '" + CStr(adc("UoAsOfDt")) + "'") + "), 0) as OdExpQty "
        
  If adc("UoYN1") = "Y" Then
    ' Zubin 213 (Cur Cnd and ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    ' ***** Manali 3.7.1 - Pending Rm Report added - ws_ExpJoin, wOrdAmdJoin, wOdWtCnd  added For Trading Reports
'      wMtchStr = " IsNull(M.OdCoCd, '') as IdCoCd, IsNull(M.OdTc, '') as IdTc, IsNull(M.OdYy, '') as IdYy, " + _
                 "IsNull(M.OdChr, '') as IdChr, IsNull(M.OdNo, 0) as IdNo, IsNull(M.OdSr, 0) as IdSr, " + _
                 "IsNull(MOm.OmDt, '') as InDt, '' as InExpNo, " + _
                 "IsNull(MOm.OmCmCd, '') as InCmCd, IsNull(M.OdJLWt, 0) as IdWt, " + _
                 "IsNull(" + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchRt/BOm.OmCnvFct) else (B.OdJLVchRt*BOm.OmCnvFct) end) ", "B.OdJLVchRt") + ", 0) As idiFOB " + _
                 "From OrdDsg B Join OrdMst BOm on BOm.OmIdNo=B.OdOmIdNo and BOm.OmCoCd= B.OdCoCd and BOm.OmTc= B.OdTc and BOm.OmYy= B.OdYy and BOm.OmChr= B.OdChr and " + _
                 "BOm.OmNo= B.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and BOm.OmPrtKey=B.OdPrtKey ", "") + _
                 "Join CustMst on CmCtg = BOm.OmCmCtg and CmCd= BOm.OmCmCd " + ws_Join + _
                 "Left Outer Join OrdDsg M on M." + wBaseRefIdKey + " =B.OdIdNo and (M.OdTc='LIN' Or M.OdCoCd = B.OdCoCd) " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "") + _
                 IIF(adc("UoAsOfDt") = "01/01/80", " ", " and M.OdOmDt <= '" + CStr(adc("UoAsOfDt")) + "' ") + _
                 "Left Outer Join OrdMst MOm on MOm.OmIdNo=M.OdOmIdNo and MOm.OmCoCd= M.OdCoCd and MOm.OmTc= M.OdTc and MOm.OmYy= M.OdYy and MOm.OmChr= M.OdChr and MOm.OmNo= M.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and MOm.OmPrtKey=M.OdPrtKey ", "") + _
                 ws_ExpJoin + " where (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= 'JT' " + IIF(gs_Partition = ctCurrPrtn, " and B.OdPrtKey='" + ctCurrPrtn + "' ", "") + wOdWtCnd
      wMtchStr = " IsNull(M.OdCoCd, '') as IdCoCd, IsNull(M.OdTc, '') as IdTc, IsNull(M.OdYy, '') as IdYy, " + _
                 "IsNull(M.OdChr, '') as IdChr, IsNull(M.OdNo, 0) as IdNo, IsNull(M.OdSr, 0) as IdSr, " + _
                 "IsNull(MOm.OmDt, '') as InDt, '' as InExpNo, " + _
                 "IsNull(MOm.OmCmCd, '') as InCmCd, IsNull(M.OdJLWt, 0) as IdWt, " + _
                 "IsNull(" + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchRt/BOm.OmCnvFct) else (B.OdJLVchRt*BOm.OmCnvFct) end) ", "B.OdJLVchRt") + ", 0) As idiFOB " + _
                 "From OrdDsg B " + _
                 "Join OrdMst BOm on BOm.OmIdNo=B.OdOmIdNo and BOm.OmCoCd= B.OdCoCd and BOm.OmTc= B.OdTc and BOm.OmYy= B.OdYy and BOm.OmChr= B.OdChr and " + _
                 "BOm.OmNo= B.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and BOm.OmPrtKey=B.OdPrtKey ", "") + _
                 "Join RmMst On RmCd=OdJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OdPrtKey ", "") + _
                 " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=B.OdJLLn1 " + _
                 "Join CustMst on CmCtg = BOm.OmCmCtg and CmCd= BOm.OmCmCd " + ws_Join + _
                 "Left Outer Join OrdDsg M on M." + wBaseRefIdKey + " =B.OdIdNo and (M.OdTc='LIN' Or M.OdCoCd = B.OdCoCd) " + IIF(gs_Partition = ctCurrPrtn, " and M.OdPrtKey=B.OdPrtKey ", "") + _
                 IIF(adc("UoAsOfDt") = "01/01/80", " ", " and M.OdOmDt <= '" + CStr(adc("UoAsOfDt")) + "' ") + _
                 "Left Outer Join OrdMst MOm on MOm.OmIdNo=M.OdOmIdNo and MOm.OmCoCd= M.OdCoCd and MOm.OmTc= M.OdTc and MOm.OmYy= M.OdYy and MOm.OmChr= M.OdChr and MOm.OmNo= M.OdNo " + IIF(gs_Partition = ctCurrPrtn, " and MOm.OmPrtKey=M.OdPrtKey ", "") + _
                 ws_ExpJoin + " where (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= 'JT' " + IIF(gs_Partition = ctCurrPrtn, " and B.OdPrtKey='" + ctCurrPrtn + "' ", "") + wOdWtCnd
   
  Else
    ' Zubin 213 (ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    wMtchStr = " '' as IdCoCd,'' as IdTc, '' as IdYy, '' as IdChr, 0 as IdNo, 0 as IdSr, '' as InDt, " + _
               "'' as InExpNo, '' as InCmCd, 0 as IdWt, 0 As idiFOB " + _
               "From OrdDsg B " + _
               "Join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
               "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
               "Join RmMst On RmCd=OdJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OdPrtKey ", "") + _
               " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=B.OdJLLn1 " + _
               "Join CustMst on CmCtg = OmCmCtg and CmCd= OmCmCd  " + _
                ws_Join + ws_ExpJoin + " where " + _
               " (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= 'JT' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + wOdWtCnd
    
'    bef Trading Module
'    wMtchStr = " '' as IdTc, '' as IdYy, '' as IdChr, 0 as IdNo, 0 as IdSr, '' as InDt, " + _
'               "'' as InExpNo, '' as InCmCd, 0 as IdQty, 0 As idiFOB " + _
'               "From OrdDsg B Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
'               "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
'               "Join CustMst on CmCtg = 'C' and CmCd= OmCmCd  " + _
'               "Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and " + _
'               "DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + ws_Join + " where " + " (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc)= 'SO' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")
  End If
 'Prev
'  If adc("UoYN3") = "Y" Then
'    wOrdDt = "OdExpDelDt"
'    Rep.hDelDt.SetText "Exp Del Dt"
'  ElseIf adc("UoYN3") = "N" Then
'    wOrdDt = "OdDelDt"
'    Rep.hDelDt.SetText "Prd Del Dt"
'  End If

'  OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd,
'  OdSfx, OdDmSz, OdOrdQty, OdSalPrc, OdKt, OdPrtCd , DmSalCtg,
'  OdExpDelDt, OdJLExpWt, qDelDt, IdTc, IdYy, IdChr, IdNo, IdSr, InDt,
'  InExpNo, InCmCd, IdQty, idiFOB

  
  '********geeta********Emr206***********Replaced + wOrdDt + " as qDelDt with OdDelDt
  ' Zubin 213
  '*** Jay 3.2.0 [PicNm]
  ' ***** Manali 3.5.0 - 11/11/08 - [OdSalRem, qDetPos, qDispSalRemYN] added
  ' ***** Manali 3.7.1 - Pending Rm Report added - wOrdWt  Used For Order Qty
    '578.22 conversion factor added to OdJLVchRt & OdJLVchVal
    wSqlStrg = "Select " + grpflds + ", " + _
            "B.OdCoCd, B.OdTc, B.OdYy, B.OdChr, B.OdNo, B.OdSr, BOm.OmDt, " + _
            "BOm.OmCmCd, BOm.OmCmCtg, BOm.OmPoNo, BOm.OmPoDt, BOm.OmCmCurCd,  " + _
            " B.OdJLRmCd, B.OdJLLOtNo, B.OdPoNo as OdPoNo, RmCtg, B.OdJLQty, B.OdJLWt," + IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchRt/BOm.OmCnvFct) else (B.OdJLVchRt*BOm.OmCnvFct) end) ", "B.OdJLVchRt") + " as OdJLVchRt, (B.OdJLWt - B.OdJLExpWt) as qBalWt, " + _
            " RmSCtg, B.OdJLLn1, B.OdJLLn2, B.OdJLLn3, " + _
            "Convert(Decimal(14,3), B.OdJLWt/(case when RmCtg in ('D', 'C') then 5 else 1 end)) as qWtInGms,   " + _
            "Convert(Decimal(14,3), B.OdJLWt/(case when RmCtg in ('D', 'C') then 5 else 1 end) " + _
            "*(case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityRt Else 0 end) )" + _
            "/(case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select RmPurityRt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= RmCtg), 1) " + _
            "else 1 end) as qPureWt, " + _
             IIF(wb_InBaseCur, " (Case when BOm.OmMulDiv= 'M' then (B.OdJLVchVal/BOm.OmCnvFct) else (B.OdJLVchVal*BOm.OmCnvFct) end) ", "B.OdJLVchVal") + " as OdJLVchVal, " + wMtchStr + _
            wAllPndCnd + wCnd + wOdWtCnd + " Order By B.OdCoCd, B.OdTc, B.OdYy, B.OdChr, B.OdNo, B.OdSr "

'    wSqlStrg = "Select " + grpflds + ", " + _
            "B.OdCoCd as OdCoCd, B.OdTc as OdTc, B.OdYy as OdYy, B.OdChr as OdChr, B.OdNo as OdNo, B.OdSr as OdSr, BOm.OmDt as OmDt, " + _
            "BOm.OmCmCd as OmCmCd, BOm.OmPoNo as OmPoNo, BOm.OmPoDt as OmPoDt,  " + _
            " B.OdJLRmCd, B.OdJLLOtNo, B.OdPoNo as OdPoNo, RmCtg, OdJLQty, OdJLVchRt, (OdJLWt - OdJLExpWt) as qBalWt, " + _
            " RmSCtg, OdJLLn1, OdJLLn2, OdJLLn3, " + _
            "Convert(Decimal(14,3), OdJLWt/(case when RmCtg in ('D', 'C') then 5 else 1 end)) as qWtInGms,   " + _
            "Convert(Decimal(14,3), OdJLWt/(case when RmCtg in ('D', 'C') then 5 else 1 end) " + _
            "*(case when RmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityRt Else 0 end) )" + _
            "/(case when RmCtg in ('G', 'P', 'S', 'L', 'M') then " + _
                          "IsNull((Select RmPurityRt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= RmCtg), 1) " + _
            "else 1 end) as qPureWt, " + _
            "OdJLVchVal " + _
            "From OrdDsg B " + _
            "Join OrdMst BOm on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and " + _
            "OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
            "Join RmMst On RmCd=OdJLRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OdPrtKey ", "") + _
            "  Join CustMst on CmCtg = OmCmCtg and CmCd= OmCmCd  " + _
            ws_Join + ws_ExpJoin + " where " + _
            " (Select PValue from Param where PTyp= 'TC' and PMCd= B.OdTc)= '" + gs_OmTcTyp + "' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + _
            wAllPndCnd + wCnd + wOdWtCnd + " Order By B.OdCoCd, B.OdTc, B.OdYy, B.OdChr, B.OdNo, B.OdSr "
  '**** Bef Trading Module
'  wSqlStrg = "Set Dateformat DMY Select " + grpflds + ", " + _
'             "OdTc, OdYy, OdChr, OdNo, OdSr, OmDt, OmCmCd, OmPoNo, DmCtg, OdDmCd, OdSfx, " + _
'             "OdDmSz, OdPicNm, OdOrdQty, " + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then (OdSalPrc/OmCnvFct) else (OdSalPrc*OmCnvFct) end) ", "OdSalPrc") + " as OdSalPrc, " + _
'             "OdKt, OdPrtCd, DmSalCtg, OdExpDelDt, " + _
'             wExpWt + ", OdDelDt as qDelDt, Replace(OdSalRem,char(13)+ char(10),' ') as OdSalRem, '" + CStr(wDetPos) + "' as qDetPos, '" + ADC("UoYN4") + "' as qDispSalRemYN, " + wMtchStr + wAllPndCnd + wCnd + _
'             IIF(wAddlCnd <> "", IIF(wCnd <> "", " And ", " Where ") + wAddlCnd, "") + wCrmCnd
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  If adc("UoYN1") = "N" Then
     Rep.BoxHd.Suppress = True
     Rep.BoxDet.Suppress = True
     Rep.BoxFt.Suppress = True
'     Rep.DetSec2.Suppress = True
   End If
 
    ' Zubin 213
    Dim ws_BaseCurCd As String
    ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                    "HCoCd= '" + adc("UoCoCdFr") + "'")
    If wb_InBaseCur Then
      Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    Else
      Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
    End If
    ' Zubin 213
  
  
    
  CRV_REP.DisplayGroupTree = False
    
    wTcDesc = moCn.GetFldVal("Select PDesc from Param Where PTyp='TC' and PMCd='" + adc("UoOmTc") + "' and PValue='JT' ")
    Rep.TxtHead.SetText "Pending " + wTcDesc + " (Loose Rm)"
    Rep.TxtAsOfDt.SetText adc("UoAsOfDt")

    
 '   ShowRepSummFld Rep.Fldr1BalQty, Rep.Fldr2BalQty, Rep.Fldr3BalQty, Rep.FldrBalQty, GRP_REP.GrpOpts, "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd"
  
'    Rep1.FldOdOrdQty.Suppress = True: Rep1.FldOdSalPrc.Suppress = True
'
'
'    Rep1.HOdDmId.Suppress = True: Rep1.HOdDmSz.Suppress = True: Rep1.HOdDmCd.Suppress = True
'    Rep1.HOdSfx.Suppress = True: Rep1.HDmCtg.Suppress = True: Rep1.hOdSalPrc.Suppress = True
'    Rep1.hOdCstPrc.Suppress = True: Rep1.hProfit.Suppress = True: Rep1.hProfitPer.Suppress = True
'    Rep1.hOdOrdQty.Suppress = True: Rep1.hOdSalVal.Suppress = True
'    Rep1.gHOdOrdQty.Suppress = True: Rep1.ghOdSalVal.Suppress = True
'    Rep1.ghProfit.Suppress = True: Rep1.ghProfPer.Suppress = True
'
'
'    GRP_REP.SetFormula Rep1, "wVal", "iif({rdo.RmQw}='Q', {rdo.OrQty}, {rdo.orWt})*{rdo.OrSalRt}"
'    GRP_REP.SetFormula Rep1, "wProfit", "0"
'    GRP_REP.SetFormula Rep1, "wProfitPer", "0"
'    GRP_REP.SetFormula Rep1, "wOdSalVal", "0"
'    GRP_REP.SetFormula Rep1, "wOdCstVal", "0"
'
'    GRP_REP.SetFormula Rep1, "g_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g1_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g2_OdOrdQty", "0"
'    GRP_REP.SetFormula Rep1, "g3_OdOrdQty", "0"
'
'    Rep1.TxtHead.SetText "Pending Rm Report "
'
'
'
'    gb_ShowPic = False
'    Rep1.PicSec.Suppress = True
    If adc("UoAsOfDt") = "01/01/80" Then
        GRP_REP.SetFormula Rep, "wBalWtAsOfDt", "{rdo.qBalWt}"
    Else
        GRP_REP.SetFormula Rep, "wBalWtAsOfDt", "{rdo.OdJLWt} - {@wExpWtAsOfDt}"
    End If
    
    Rep.TxtCoCdFr.SetText adc("UoCoCdFr")
    Rep.TxtCoCdTo.SetText adc("UoCoCdTo")
    Rep.TxtCoCdSel.SetText adc("UoCoCdSel")
    
    '578.19 calling function to change Rmcode font size
    GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OdJLRmCd}", 8)
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

End Sub
'''' std code not to be changed
Private Sub ADC_Load()
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
    gb_CoCdFor = False
    If ws_MultiCoMod = True Then
      Set adc.FirNKeyCtl = adc("UoCoCdFr")
    Else
      Set adc.FirNKeyCtl = adc("UoOmTcFr")
    End If
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    
    'Set adc.FirNKeyCtl = adc("UoOmTcFr") 'adc("UoCoCdFr")
    
    Call InitProp(Me)
     
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
    If ws_MultiCoMod = True Then
        adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
    Else
        adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
    End If
    ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
     
'    ' ***** Manali Trading Module
      
      gs_PTyp = "TC"
      gs_OmTcTyp = ctJtTcTyp
      gs_CmCtg = "C"
      
      '**** Manali Trading Module
    'gb_CmCtgFor = False ' MANALI 3.8.0
    gs_CmCtg1 = ""
    
    Call SetGroupSort
    TAB_REP.Tab = 0
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  gb_CmCtgFor = False ' MANALI 3.8.0
  gs_CmCtg1 = ""
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  '''gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  gb_CmCtgFor = False ' MANALI 3.8.0
    gs_PTyp = "TC"
    gs_OmTcTyp = ctJtTcTyp
    gs_CmCtg = moCn.GetFldVal("Select TmCmCtg from TcMst Where TmCoCd='" + gs_CoCd + "' and TmTc='" + adc("UoOmTc") + "' ")
    gs_CmCtg = IIF(gs_CmCtg = "", "C", gs_CmCtg)
    
  
  gs_CmCtg1 = ""
  SetProp Me, IdName, When
End Sub


Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  gb_CoCdFor = False  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  Select Case UCase(IdName)
  Case Is = UCase("UoOmTcFr")
      adc("UoOmTcFr") = adc("UoOmTc")
      adc("UoOmTcTo") = adc("UoOmTc")
  End Select
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPnd"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
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
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  adc("UoYN2") = "N"
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked


End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + IIF(ws_MultiCoMod = True, gs_CoCd, adc("UOCOCDFR")) + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")

'  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd = '" + ctSelfCmCd + "' ")

  
  Rep.wCoCd.SetText ws_HName
  'Rep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
   If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdFr.SetText adc("UOCOCDFR"): Rep.TxtCoCdTo.SetText adc("UOCOCDTO"): Rep.TxtCoCdSel.SetText adc("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  ' ***** Manali 3.10.0 - 12/03/12 - Multi Copmpany Allowed
  

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  'Rep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
End Sub
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  'fra_keys.Picture = Me.Picture
  'fra_keys.BorderStyle = 0
  'fra_nkeys.Picture = Me.Picture
  'fra_nkeys.BorderStyle = 0
  
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  Call SetObjColors(Me)
  '*** Jenny Colour
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

