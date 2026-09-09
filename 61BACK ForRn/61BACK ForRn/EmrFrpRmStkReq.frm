VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4C3DB8FE-9A14-4B68-B881-D13A3BD0781B}#1.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpRmStkReq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material Stock Requirement"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.Frame FraSC 
      Height          =   2085
      Left            =   3450
      TabIndex        =   110
      Top             =   5010
      Visible         =   0   'False
      Width           =   5205
      Begin VB.TextBox TxtPwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2730
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   63
         Top             =   1125
         Width           =   2340
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   38
         Left            =   2730
         TabIndex        =   62
         ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
         Top             =   840
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         DataType        =   4
         DataField       =   "UOYN3"
         ReCalcParent    =   "UOYN1"
         IdName          =   "UOYN3"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   50
         Left            =   2730
         TabIndex        =   64
         ToolTipText     =   "Enter 'Y' For Cost Value, 'N' for Sale Value or Blank For No Value."
         Top             =   1410
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         DataField       =   "UOYN"
         IdName          =   "UOYN"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   58
         Left            =   2730
         TabIndex        =   65
         ToolTipText     =   "Enter 'Y' For Calculated Pointer, 'N' for Stored Actual Pointer"
         Top             =   1695
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   3
         DataField       =   "UOYN1"
         ReCalcOn        =   "UOYN3"
         IdName          =   "UOYN1"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   62
         Left            =   2730
         TabIndex        =   60
         ToolTipText     =   "Show Order Box Details (Y/N)"
         Top             =   270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         DataType        =   4
         DataField       =   "UOYN2"
         ReCalcParent    =   "UOYN1"
         IdName          =   "UOYN2"
      End
      Begin MwfCtl.MWCTL_MED atxt 
         Height          =   285
         Index           =   65
         Left            =   2730
         TabIndex        =   61
         ToolTipText     =   "Show Stock Box Details (Y/N)"
         Top             =   555
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         DataType        =   4
         DataField       =   "UOYN4"
         ReCalcParent    =   "UOYN1"
         IdName          =   "UOYN4"
      End
      Begin VB.Label Label4 
         Caption         =   "Show Stock Box Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   120
         Top             =   555
         Width           =   2505
      End
      Begin VB.Label Label3 
         Caption         =   "Show Order Box Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   116
         Top             =   270
         Width           =   2500
      End
      Begin VB.Label Label2 
         Caption         =   "Show Prd Rm Qty && Wt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   114
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label ALBL 
         Caption         =   "Show Rm Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   90
         TabIndex        =   113
         Top             =   1410
         Width           =   2175
      End
      Begin VB.Label ALBL 
         Caption         =   "Enter SEO Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   90
         TabIndex        =   112
         Top             =   1125
         Width           =   2265
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Calculated Pointer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   111
         Top             =   1695
         Width           =   2745
      End
   End
   Begin VB.CommandButton CmdSc 
      Caption         =   "&Extra Scope"
      Height          =   390
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Analysis Scope Options"
      Top             =   7200
      Width           =   1125
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10065
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   3210
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7125
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4065
      TabIndex        =   76
      Top             =   7185
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   77
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
         TabIndex        =   78
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
      Left            =   30
      TabIndex        =   68
      Top             =   7140
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   979
      AllowDelete     =   0   'False
      Automate        =   -1  'True
      EntType         =   1
      BtnPicture      =   "EmrFrpRmStkReq.frx":0000
      BtnWidth        =   885
      BtnHeight       =   435
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   5280
      TabIndex        =   75
      Top             =   7155
      Visible         =   0   'False
      Width           =   735
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   120
         TabIndex        =   72
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
      BackColor       =   &H00C0C0C0&
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
      Height          =   7110
      Left            =   0
      TabIndex        =   71
      Top             =   60
      Width           =   11835
      Begin TabDlg.SSTab TAB_REP 
         Height          =   7020
         Left            =   60
         TabIndex        =   70
         Top             =   60
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   12383
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRmStkReq.frx":0B1E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpRmStkReq.frx":0B3A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Enabled         =   0   'False
            Height          =   4005
            Index           =   1
            Left            =   -74910
            TabIndex        =   74
            Top             =   450
            Width           =   8925
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   69
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   6615
            Index           =   0
            Left            =   60
            TabIndex        =   73
            Top             =   360
            Width           =   11640
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4140
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1365
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "OrRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6840
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1365
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   795
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
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   795
               Width           =   4800
               _ExtentX        =   8467
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
               Top             =   225
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "OrRmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1485
               TabIndex        =   5
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   795
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
               TabIndex        =   11
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1365
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "OrRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   4140
               TabIndex        =   15
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   1650
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
               TabIndex        =   14
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   1650
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
               TabIndex        =   17
               ToolTipText     =   "Enter To Raw Material Breadth"
               Top             =   1935
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
               TabIndex        =   16
               ToolTipText     =   "Enter From Raw Material Breadth"
               Top             =   1935
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
               TabIndex        =   19
               ToolTipText     =   "Enter To Sale Rate"
               Top             =   2220
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
               TabIndex        =   18
               ToolTipText     =   "Enter From Sale Rate"
               Top             =   2220
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
               TabIndex        =   21
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   2505
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
               TabIndex        =   20
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   2505
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
               TabIndex        =   23
               ToolTipText     =   "Enter To Production Quantity"
               Top             =   2790
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
               TabIndex        =   22
               ToolTipText     =   "Enter From Production Quantity"
               Top             =   2790
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
               TabIndex        =   25
               ToolTipText     =   "Enter To Balance Production"
               Top             =   3075
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
               TabIndex        =   24
               ToolTipText     =   "Enter From Balance Production"
               Top             =   3075
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
               TabIndex        =   37
               ToolTipText     =   "Enter To Design Category"
               Top             =   3930
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
               TabIndex        =   38
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   3930
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   36
               ToolTipText     =   "Enter From Design Category"
               Top             =   3930
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
               TabIndex        =   40
               ToolTipText     =   "Enter To Design Code"
               Top             =   4215
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
               TabIndex        =   41
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   4215
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   39
               ToolTipText     =   "Enter From Design Code"
               Top             =   4215
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
               TabIndex        =   43
               ToolTipText     =   "Enter To Design Size"
               Top             =   4500
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
               TabIndex        =   42
               ToolTipText     =   "Enter From Design Size"
               Top             =   4500
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
               Left            =   4140
               TabIndex        =   45
               ToolTipText     =   "Enter To Suffix"
               Top             =   4785
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
               Left            =   1485
               TabIndex        =   44
               ToolTipText     =   "Enter From Suffix"
               Top             =   4785
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
               TabIndex        =   47
               ToolTipText     =   "Enter To Customer"
               Top             =   5070
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
               TabIndex        =   48
               ToolTipText     =   "Enter Customer Selection"
               Top             =   5070
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   46
               ToolTipText     =   "Enter From Customer"
               Top             =   5070
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
               TabIndex        =   50
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   5355
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
               TabIndex        =   51
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   5355
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   49
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   5355
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
               TabIndex        =   53
               ToolTipText     =   "Enter To Order Date"
               Top             =   5640
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
               TabIndex        =   52
               ToolTipText     =   "Enter From Order Date"
               Top             =   5640
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
               TabIndex        =   55
               ToolTipText     =   "Enter To Prd Delivery Date"
               Top             =   5925
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
               TabIndex        =   54
               ToolTipText     =   "Enter From Prd Delivery Date"
               Top             =   5925
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
               TabIndex        =   30
               ToolTipText     =   "Enter To Order Year"
               Top             =   3360
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
               CmpStr          =   "OrYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2265
               TabIndex        =   27
               ToolTipText     =   "Enter From Order Year"
               Top             =   3360
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
               CmpStr          =   "OrYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1485
               TabIndex        =   26
               ToolTipText     =   "Enter From Order Tc"
               Top             =   3360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OrTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2685
               TabIndex        =   28
               ToolTipText     =   "Enter From Order Character"
               Top             =   3360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
               CmpStr          =   "OrChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4140
               TabIndex        =   29
               ToolTipText     =   "Enter To Order Tc"
               Top             =   3360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OrTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5325
               TabIndex        =   31
               ToolTipText     =   "Enter To Order Character"
               Top             =   3360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
               CmpStr          =   "OrChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5025
               TabIndex        =   35
               ToolTipText     =   "Enter To Order Serial"
               Top             =   3645
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "OrSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2385
               TabIndex        =   33
               ToolTipText     =   "Enter From Order Serial"
               Top             =   3645
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "OrSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1485
               TabIndex        =   32
               ToolTipText     =   "Enter From Order Number"
               Top             =   3645
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
               CmpStr          =   "OrNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4140
               TabIndex        =   34
               ToolTipText     =   "Enter To Order Number"
               Top             =   3645
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
               CmpStr          =   "OrNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   4140
               TabIndex        =   9
               ToolTipText     =   "Enter Raw Material Sub Shape To Range"
               Top             =   1080
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOSUBSHPTO"
               IdName          =   "UOSUBSHPTO"
               CmpStr          =   "OrSubShp <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   6840
               TabIndex        =   10
               ToolTipText     =   "Enter Raw Material Sub Shape Selection"
               Top             =   1080
               Width           =   4800
               _ExtentX        =   8467
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
               TabIndex        =   8
               ToolTipText     =   "Enter Raw Material Sub Shape From Range"
               Top             =   1080
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOSUBSHPFR"
               IdName          =   "UOSUBSHPFR"
               CmpStr          =   "OrSubShp >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1485
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   225
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OrCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   1485
               TabIndex        =   56
               ToolTipText     =   "Enter From  Priority Code"
               Top             =   6210
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
               TabIndex        =   57
               ToolTipText     =   "Enter To  Priority Code"
               Top             =   6210
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
               TabIndex        =   58
               ToolTipText     =   "Enter  Priority Code Selection"
               Top             =   6210
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd  In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4140
               TabIndex        =   3
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   6840
               TabIndex        =   4
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   510
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   61
               Left            =   1485
               TabIndex        =   2
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   10065
               TabIndex        =   117
               ToolTipText     =   "Enter To YearMonth"
               Top             =   2790
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMTO"
               IdName          =   "UOYYMMTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   8025
               TabIndex        =   118
               ToolTipText     =   "Enter From YearMonth"
               Top             =   2790
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMFR"
               IdName          =   "UOYYMMFR"
            End
            Begin VB.Label ALBL 
               Caption         =   "YyyyMm"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   6840
               TabIndex        =   119
               Top             =   2790
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Location"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   115
               Top             =   510
               Width           =   1425
            End
            Begin VB.Label Label1 
               Caption         =   "Priority Cd"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   109
               Top             =   6210
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   108
               Top             =   240
               Width           =   105
            End
            Begin VB.Label ALBL 
               Caption         =   "Sub Shp"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   107
               Top             =   1080
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Ord Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   106
               Top             =   3360
               Width           =   1725
            End
            Begin VB.Label ALBL 
               Caption         =   "Ord No/Sr"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   105
               Top             =   3645
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   104
               Top             =   3360
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   103
               Top             =   3360
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   102
               Top             =   3360
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   101
               Top             =   3360
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   100
               Top             =   3645
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackColor       =   &H00C0C0C0&
               Caption         =   "/"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   99
               Top             =   3645
               Width           =   105
            End
            Begin VB.Label ALBL 
               Caption         =   "Prd Del Dt"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   98
               Top             =   5925
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Ord Dt"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   97
               Top             =   5640
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "PO No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   96
               Top             =   5355
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Customer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   95
               Top             =   5070
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Suffix"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   94
               Top             =   4785
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Design Size"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   93
               Top             =   4500
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Design Cd"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   92
               Top             =   4215
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Design Ctg"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   91
               Top             =   3930
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Bal Prd"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   90
               Top             =   3075
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Prd Qty"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   89
               Top             =   2790
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Cost Rate"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   88
               Top             =   2505
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Sale Rate"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   87
               Top             =   2220
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Breadth"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   86
               Top             =   1935
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Len/Sv"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   85
               Top             =   1650
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Co Cd/Rm Ctg"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   84
               ToolTipText     =   "Location"
               Top             =   225
               Width           =   1485
            End
            Begin VB.Label ALBL 
               Caption         =   "Rm SubCtg"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   83
               Top             =   795
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Rm Code"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               Left            =   0
               TabIndex        =   82
               Top             =   1365
               Width           =   1425
            End
            Begin VB.Label ALBL 
               Caption         =   "Selection"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   81
               Top             =   15
               Width           =   1575
            End
            Begin VB.Label ALBL 
               Caption         =   "To"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   80
               Top             =   15
               Width           =   1275
            End
            Begin VB.Label ALBL 
               Caption         =   "From"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   79
               Top             =   15
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRmStkReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REP As New EmrRepRmStkReq
Dim moCn As MwfLib.MDOConnection
Dim mf_CalcPtr As Boolean
Private Sub SetGroupSort()
'''''  GRP_REP.Add "Rm Code", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)+Rtrim(OrRmCd)", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hOrRmCtg,hOrRmSCtg,hOrRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "

  'GRP_REP.Add "Rm Ctg", Array("OrRmCtg", "Max(RmCtg)"), "",  Array("OrRmCtg", "Max(RmCtg)"), "hOrRmCtg", "", "", "", Array("(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) ", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = Max(RmCtg))")
  GRP_REP.Add "Rm Ctg", Array("OrRmCtg", "Max(RmCtg)"), "", "OrRmCtg", "hOrRmCtg", "", "", "", Array("(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) ", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = Max(RmCtg))")
  GRP_REP.Add "Rm SubCtg", Array("OrRmCtg+OrRmSCtg", "max(RmCtg)+max(RmSCtg)"), _
     Array("'('+ OrRmCtg+ ') '+ OrRmSCtg", "'('+ max(RmCtg)+ ') '+ max(RmSCtg)"), _
     "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg", "", "", _
    Array("(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) ", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(RmCtg) and PSCd= max(RmSCtg)) ")
  
  GRP_REP.Add "Rm Code", Array("OrRmCtg+OrRmSCtg+OrRmCd", "Max(RmCtg)+Max(RmSCtg)+Max(RmCd)"), Array("'('+OrRmCtg+'/'+OrRmSCtg+') '+OrRmCd", "'('+Max(RmCtg)+'/'+Max(RmSCtg)+') '+Max(RmCd)"), "OrRmCtg,OrRmSCtg,OrRmCd", "hOrRmCtg,hOrRmSCtg,hOrRmCd", _
    "Rm Ctg,Rm SubCtg", "", "", _
    Array("SubString(RmDesc, 1, 30) ", "SubString(max(RmDesc), 1, 30) ")
    'Array("(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCd) ", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = Max(RmCd))")

'''''  GRP_REP.Add "Sub Shp", Array("OrSubShp", "max(OrSubShp)"), "", "", "", "", "", "", Array("(Select PDesc From Param Where PTyp= 'SUBSHP' and PMCd = OrSubShp) ", "(Select PDesc From Param Where PTyp= 'SUBSHP' and PMCd = max(OrSubShp)) ")
  
  GRP_REP.Add "Parent Cd", _
        Array("OrRmCtg+OrRmSCtg+(case when RmParentCd= '' then OrRmCd else RmParentCd end)", _
        "max(RmCtg)+max(RmSCtg)+(case when max(RmParentCd)= '' then max(RmCd) else max(RmParentCd) end)"), _
        Array("'('+ OrRmCtg+'/'+ OrRmSCtg+ ') '+ (case when RmParentCd= '' then OrRmCd else RmParentCd end)", _
         "'('+ max(RmCtg)+'/'+ max(RmSCtg)+ ') '+ (case when max(RmParentCd)= '' then max(RmCd) else max(RmParentCd) end)"), _
      "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg,Rm SubCtg,BaseCd + RmCd", "", "", _
       Array("IsNull((Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd),'') ", "IsNull((Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd)),'') ")
  
  GRP_REP.Add "Range Size", Array("OrRmCtg+OrRmSCtg+IsNull(RrRngSz, '')", "Max(RmCtg)+Max(RmSCtg)+Max(IsNull(RSZ.RrRngSz, ''))"), Array("'('+OrRmCtg+', '+OrRmSCtg+') '+IsNull(RrRngSz, '')", "'('+Max(RmCtg)+', '+Max(RmSCtg)+') '+Max(IsNull(RSZ.RrRngSz, '') )"), "qRrRngSz", "hRrRngSz", "Rm Ctg,Rm SubCtg", "", "", ""

  '###
  'GRP_REP.Add "Rm Len/Sv", "Str(OrLn1,6,3)", "", "OrLn1", "hOrL  n1", "", "", "", ""
  'GRP_REP.Add "Rm Len 2", "str(OrLn1,6,3)+' '+str(OrLn1,6,3)", "'('+ LTrim(str(OrLn1,6,3))+ ') '+ LTrim(str(OrLn2,6,3))", "OrLn1,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
  
  'GRP_REP.Add "Rm Len/Sv", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else Str(OrLn1,6,3) end)", "", "wLnSv", "hOrLn1", "", "", "", ""
   '****************Geeta**************Emr206*********12/04/04
'''''  Dim wsPtrYN As String
'''''  wsPtrYN = "PValue3 "
'''''  GRP_REP.Add "Rm Len/Sv", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else (case when " + wsPtrYN + "='Y' then Str(OrLn1,7,4) else Str(OrLn1,6,3) end) end)", "", "wLnSv", "hOrLn1", "", "", "", ""
  
'''''  GRP_REP.Add "Rm Len 2", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else Str(OrLn1,6,3) end)+' '+str(OrLn2,6,3)", "'('+ (case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else LTrim(Str(OrLn1,6,3)) end)+ ') '+ LTrim(str(OrLn2,6,3))", "wLnSv,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
  
  'Manoj 2.10.0
  'If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", "Str(OrRmPtr,7,4)", "", "OrRmPtr", "hOrRmPtr", "", "", "", ""
'''''  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", Array("Str(OrRmPtr,7,4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') And PValue3 = 'Y' Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End,7,4)"), Array("Str(OrRmPtr, 7, 4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') And PValue3 = 'Y' Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End, 7, 4)"), "OrRmPtr", "hOrRmPtr", "", "", "", ""
  'Manoj 2.10.0
  
'''''  GRP_REP.Add "RmCd Sale Rate", "str(OrSalRt,9,2)", "", "OrSalRt", "hOrSalRt", "RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
'''''  GRP_REP.Add "RmCd Cost Rate", "str(OrCstRt,9,2)", "", "OrCstRt", "hOrCstRt", "RmCd Sale Rate,(RmCd) SaleRt,(RmCd) CostRt", "", "", ""
'''''  GRP_REP.Add "(RmCd) SaleRt", "RmCd+str(OrSalRt,9,2)", "'('+RmCd+') '+LTrim(str(OrSalRt,9,2))", "OrRmCtg,OrRmSCtg,OrRmCd,OrSalRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrSalRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) CostRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
'''''  GRP_REP.Add "(RmCd) CostRt", "RmCd+str(OrCstRt,9,2)", "'('+RmCd+') '+LTrim(str(OrCstRt,9,2))", "OrRmCtg,OrRmSCtg,OrRmCd,OrCstRt", "hOrRmCtg,hOrRmSCtg,hOrRmCd,hOrCstRt", "Rm Code,Rm Ctg,Rm SubCtg,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt", "", "", " '('+ RTrim(RmCtg)+ '/'+ RmSCtg+ ')'"
  'GRP_REP.Add "RmCd Sale Rate", "OrSalRt", "", "OrSalRt", "hOrSalRt", "OrCstRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "
  'GRP_REP.Add "RmCd Cost Rate", "OrCstRt", "", "OrCstRt", "hOrCstRt", "OrSalRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "
  '###
  'GRP_REP.Add "(RmCd) Len/Sv", "OrRmCd+Str(OrLn1,4,2)", "'('+OrRmCd+') '+LTrim(str(OrLn1,4,2))", "OrRmCd,OrLn1", "hOrRmCd,hOrLn1", "Rm Ctg,Rm SubCtg,Rm Code", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  'GRP_REP.Add "(RmCd) Len 2", "OrRmCd+str(OrLn1,4,2)+' '+str(OrLn1,4,2)", "OrRmCd+' ('+ LTrim(str(OrLn1,4,2))+ ') '+ LTrim(str(OrLn2,4,2))", "OrRmCd,OrLn1,OrLn2", "hOrRmCd,hOrLn1,hOrLn2", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Len/Sv", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  
  'GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd,OmDt,OmPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo", "Order Dt,Customer Code,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
'''''  GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd", "hOrdNo,hOmCmCd", "Order Dt,Customer Code,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
'''''  GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,(DsgCd)Colour", "", "", "(Select DmDesc From DsgMst Where DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '') "
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
'''''  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.Text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
'''''  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'PRTCD' and PmCd= OdPrtCd) "
'''''  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "wOrdNo,OmCmCd,OmDt,OmPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo", "Order Dt,Customer Code,OrderNo", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
'''''  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
'''''  GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "OdDmCd", "hOdDmCd", "Design Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
'''''  GRP_REP.Add "(OrdKt)Colour", "OdKt+OdDmCol", "'('+OdKt+') '+OdDmCol", "OdKt,OdDmCol", "", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt)+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
  
End Sub
Private Sub adc_setreprecsource()
'  *** Report Sql ***
'  Select Space(1) as qTag, space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  OrRmCd, OrRmCtg, OrRmSCtg, PValue3 as qPtrYN, OrLn1, OrLn2, OrRmPtr, OrQty,
'  OrWt as qOrCarats,  OrWt as qOrWt, OrWt as qOrPureWt,
'  OrSalRt, OrCstRt, OdPrdQty, OdPrdQty as qOdBalPrd,
'  OrTc, OrYy,OrChr,OrNo,OrSr, OrSrNo, OdDmCd, OdDmSz, OdSfx,
'  OmCmCd , OmPoNo, DmCtg, OmDt, OdDelDt,RrRngSz as qRrRngSz ,SLoc As qSLoc, SRmCrWt As qStkWt, SRmCrQty As qStkQty /*,
'  SRmStkRt as qSRmStkRt*/
'  from OrdRm, OrdDsg, OrdMst, DsgMst,Param,SYyMm ,RmRt where 1=2
  '*** Report Sql ***
  
' ********************************************** Zubin 211 ********************************************** '
' **** PN: Whenever SQL of Rm Requirement (static) report changes, change the SQL of this report also **** '
' ****          RrRngSz Should Be Considered in the Query          **** '
' ********************************************** Zubin 211 ********************************************** '
  
  Dim grpflds0 As String, grpflds1 As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wPurityWt As String
  Dim ws_RmQty As String, ws_RmWt As String, ws_RmPtr As String
  
  ' **** Zubin 211 **** '
  Dim ws_OrdBy As String
  Dim ws_SLocCnd As String
  ' **** Zubin 211 **** '
  
  If FraSC.Visible = True Then CmdSC_Click
  
  Set REP = Nothing
  Set adc.RepSource = REP
  Call DispCoNm
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
    
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, wi_Ctr As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For wi_Ctr = 1 To 6
    If wGrp(wi_Ctr) = "(Detail)" Then
      wDetPos = wi_Ctr: Exit For
    End If
  Next wi_Ctr
  
  '### Note TdByLoc,TdByWrk,qRmSCtg, qRmCd fields place in G4-DET, will come from the last record of the Bag. OK.
    
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    REP.BOXORDHD.Suppress = True: REP.BOXORDDET.Suppress = True: REP.BOXORDDETFT.Suppress = True
    REP.BOXFT1.Suppress = True
    REP.BOXSTKHD.Suppress = True: REP.BOXSTKDET.Suppress = True: REP.BOXSTKDETFT.Suppress = True
    REP.BOXFT2.Suppress = True
  End If
  
  If adc("UoYN3") = "Y" Then
    ws_RmQty = "OrPrdQty as OrQty": ws_RmWt = "OrPrdWt"
  Else
    ws_RmQty = "OrQty": ws_RmWt = "OrWt"
  End If
  'Manoj 2.10.0
  If adc("UoYN1") = "Y" And adc("UoYN3") = "Y" Then
    ws_RmPtr = " Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') And PValue3 = 'Y' Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End As OrRmPtr "
'ws_RmPtr = "Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') And PValue3 = 'Y' Then " + _
"  (Case when PValue2 = 'Y' then " + _
"    OrRmPtr " + _
"  Else " + _
"    OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) End) " + _
"  Else 0 End) " + _
"End As OrRmPtr "

    
'''''    grpflds = GRP_REP.GrpFldLst(1)
  Else
    ws_RmPtr = " OrRmPtr "
'''''    grpflds = GRP_REP.GrpFldLst(0)
  End If
  'Manoj 2.10.0
  
  If adc("UoYn") = "N" Then
    GRP_REP.SetFormula REP, "wRmVal", "{rdo.OdPrdQty}*{rdo.OrSalRt}"
    GRP_REP.SetFormula REP, "g1_wValDisp", "'Rm SalVal: ' + ToText ({@g1_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g2_wValDisp", "'Rm SalVal: ' + ToText ({@g2_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g3_wValDisp", "'Rm SalVal: ' + ToText ({@g3_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g_wValDisp", "'Rm SalVal: ' + ToText ({@g_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g4_wValDisp", "'Rm SalVal: ' + ToText ({@g4_wRmVal}, 2)"
  ElseIf adc("UOYN") = "Y" Then
    GRP_REP.SetFormula REP, "wRmVal", "{rdo.OdPrdQty}*{rdo.OrCstRt}"
    GRP_REP.SetFormula REP, "g1_wValDisp", "'Rm CstVal: ' + ToText ({@g1_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g2_wValDisp", "'Rm CstVal: ' + ToText ({@g2_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g3_wValDisp", "'Rm CstVal: ' + ToText ({@g3_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g_wValDisp", "'Rm CstVal: ' + ToText ({@g_wRmVal}, 2)"
    GRP_REP.SetFormula REP, "g4_wValDisp", "'Rm CstVal: ' + ToText ({@g4_wRmVal}, 2)"
  ElseIf adc("UOYN") = "" Then
    GRP_REP.SetFormula REP, "wRmVal", "0"
    GRP_REP.SetFormula REP, "g1_wValDisp", ""
    GRP_REP.SetFormula REP, "g2_wValDisp", ""
    GRP_REP.SetFormula REP, "g3_wValDisp", ""
    GRP_REP.SetFormula REP, "g_wValDisp", ""
    GRP_REP.SetFormula REP, "g4_wValDisp", ""
  End If
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " And " + wrepcnd, "")
  
  
  wPurityWt = " /(case when OrRmCtg in ('G', 'P') then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst " + _
                          "where RmCd= PValue and PTyp= 'RMCTG' and PMCd= OrRmCtg), 1) " + _
                "else 1 end) "
   '****************Geeta**************Emr206*********12/04/04
   'qPtrYN  field included in the query

  Call TmpTbl
  wSqlStrg = "Insert Into #TmpRmReq Select IsNull(RrRngSz, ''), " + grpflds0 + ", " + _
               "OrRmCd, OrRmCtg, OrRmSCtg, PValue3 as qPtrYN, OrLn1, OrLn2, " + ws_RmPtr + ", " + ws_RmQty + ", " + _
               "(Case when (OrRmCtg= 'D' or OrRmCtg= 'C') then " + ws_RmWt + " else 0.0 end) as qOrCarats, " + _
               "(Case when (OrRmCtg= 'G' or OrRmCtg= 'P' or OrRmCtg= 'M' or OrRmCtg= 'X') then " + ws_RmWt + " else 0.0 end) as qOrWt, " + _
               "(Case when (OrRmCtg='G' or OrRmCtg='P') then (Select RmPurityWt from RmMst where RmCd=OrRmCd)" + wPurityWt + " else 0.0 end)* " + ws_RmWt + " as qOrPureWt, " + _
               "OrSalVal as OrSalRt, OrCstVal as OrCstRt, OdPrdQty, (OdPrdQty - isnull((Select sum(BQty) from Bag where " + _
               "BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + _
               "and ((BCls= 'Y') or (  exists(Select 'x' from Txnd where TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo " + _
               "and TdRmCd= OrRmCd) and isnull((select sum(TdRmWt*( " + _
                                              "Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
                                              "Else 0 end)) " + _
               "from Txnd where TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) " + _
               ", 0.0)> 0)  )), 0.0)) as qOdBalPrd, OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo, OdDmCd, " + _
               "OdDmSz, OdSfx, OmCmCd, OmPoNo, DmCtg, OmDt, OdDelDt  " + _
               "From OrdRm Join OrdDsg on OdCoCd= OrCoCd and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + _
               "Join OrdMst on OmCoCd= OrCoCd and OmTc= OrTc and OmYy= OrYy and OmChr= OrChr and OmNo= OrNo " + _
               "Join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' join RmMst on RmCd= OrRmCd " + _
               "Join Param On PTyp = 'RMSCTG' ANd PMCd = RmCtg And PSCd = RmSCtg "
  
  wSqlStrg = wSqlStrg + "Left Outer Join RmRt On RrTcTyp= 'RSZ' and RrCmCtg= 'C' and " + _
                "   RrCmCd= '" + ctSelfCmCd + "' and RrCtg= OrRmCtg and RrSCtg= OrRmSCtg and " + _
                "RrCmCurCd= '' and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' and " + _
                "RrFrLn<= (Case when PValue3= 'Y' then OrRmPtr else OrLn1 End) and " + _
                "RrToLn>= (Case when PValue3= 'Y' then OrRmPtr else OrLn1 End) " + _
                " Where (OdPrdQty - isnull((Select sum(BQty) from Bag where " + _
                "BCoCd= OrCoCd and BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr and BOdNo= OrNo and BOdSr= OrSr " + _
                "and ((BCls= 'Y') or (  exists(Select 'x' from Txnd where TdCoCd= BCoCd and TdBYy= BYy and " + _
                "TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) and " + _
                "isnull((select sum(TdRmWt*( " + _
                                              "Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
                                              "Else 0 end)) " + _
               "from Txnd where TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) " + _
               ", 0.0)> 0)  )), 0.0))>0  " + wCnd + " "

  moCn.Execute (wSqlStrg)
'Debug.Print wSqlStrg
    
'''''  wSqlStrg = " select 'O' as qTag, * , '' qRrRngSz, '' as qSLoc, 0 as qStkWt, 0 as qStkQty /*, 0 as qSRmStkRt */ From #TmpRmReq "
  
  wSqlStrg = "Select 'S' as qTag ," + grpflds1 + ", " + _
              "SRmCd as OrRmCd, Max(RmCtg) as OrRmCtg, Max(RmSCtg) OrRmSCtg, Max(qPtrYN) as qPtrYN, SRmSz as OrLn1, " + _
              "0 as OrLn2, (case when Max(qPtrYN)='Y' Then SRmSz Else 0 End ) as OrRmPtr, " + _
              "0 as OrQty, 0 as qOrCarats, 0.0 as qOrWt, " + _
              "(Case when RmCtg in ('G', 'P') then Max(RmPurityWt) when RmCtg= 'X' " + _
              " then Max(RmBasePurityWt) else 0 end) /(case when Rm.RmCtg in ('G', 'P') then " + _
              "IsNull((Select Max(RmPurityWt) from Param join RmMst on RmCd= PValue  " + _
              "where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) when Rm.RmCtg= 'X' " + _
              "then IsNull((Select RmPurityWt from Param " + _
              "join RmMst on RmCd= PValue where PTyp= 'RMCTG' and " + _
              " PMCd= IsNull((Select RmCtg From RmMst where RmCd= Max(Rm.RmBaseCd)), " + _
              " '')), 1) else 1 end) as qOrPureWt, 0 as OrSalRt, 0 as OrCstRt, " + _
              "0 as OdPrdQty, 0 as qOdBalPrd,'' OrTc ,'' OrYy,'' as OrChr,0 as  OrNo, 0 as  OrSr," + _
              "0 as OrSrNo,'' as  OdDmCd,'' as OdDmSz,'' as  OdSfx,'' as  OmCmCd," + _
              "'' as  OmPoNo,'' as  DmCtg,'01/01/80' as  OmDt,'01/01/80' as  OdDelDt, Max(IsNull(RSZ.RrRngSz, '')) qRrRngSz, " + _
              "SLoc as qSLoc, Sum(SRmDrWt- SRmCrWt) as qStkWt, Sum(SRmDrQty- SRmCrQty) as qStkQty /*, " + _
              "SRmStkRt as qSRmStkRt */" + _
              "From SYyMm Join RmMst Rm On RmCd= SRmCd " + _
              "Join Loc on LocCoCd= SCoCd and LocCd= SLoc " + _
              "Left Outer Join RmRt RSZ on RrTcTyp= 'RSZ' and RrCmCtg= 'C' and RrCmCd= '" + ctSelfCmCd + "' and " + _
              "RrCtg= RmCtg and RrSCtg= RmSCtg and RrCmCurCd= '' and RrSTWGrp= '' and RrCd= '' " + _
              "and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' and RrFrLn<= SRmSz and RrToLn>= SRmSz "
  
  If adc("UOISSLOCFR") <> "" Then ws_SLocCnd = " and LocCd >= '" + adc("UOISSLOCFR") + "' "
  If adc("UOISSLOCTO") <> "" Then ws_SLocCnd = ws_SLocCnd + " and LocCd <= '" + adc("UOISSLOCTO") + "' "
  If adc("UOISSLOCSEL") <> "" Then ws_SLocCnd = ws_SLocCnd + " and LocCd In (" + adc("UOISSLOCSEL") + ") "

  wSqlStrg = wSqlStrg + "Join (Select Distinct OrRmCd, qPtrYN, /*OrLn1, OrRmPtr, */ RrRngSz " + _
                          "From  #TmpRmReq ) ORm On SCoCd='" + adc("UOCoCdFr") + "' and OrRmCd=SRmCd and ORm.RrRngSz= IsNull(RSZ.RrRngSz, '') /* and " + _
                          " ( Case when qPtrYN='Y' then OrRmPtr else OrLn1 end )= SRmSz */ " + _
                        "Where SCoCd= '" + adc("UoCoCdFr") + "' and SYyMm>= " + CStr(adc("UoYyMmFr")) + " and SYyMm<= " + CStr(adc("UoYyMmTo")) + " and " + _
                        "LocTyp='S' and ((SRmDrWt <> 0 or SRmCrWt <> 0) /* or (SRmDrQty <> 0 Or SRmCrQty <> 0) */) " + _
                        ws_SLocCnd + _
                        "Group By SCoCd, SLoc, RmCtg, SRmCd, SRmSz /*, SRmStkRt */ "

  wSqlStrg = wSqlStrg + " Union All "
  
  wSqlStrg = wSqlStrg + " select 'O' as qTag, Grp1, Grp2, Grp3, DGrp1, DGrp2, DGrp3, G1Desc, G2Desc, G3Desc, " + _
                "OrRmCd, OrRmCtg, OrRmSCtg, qPtrYN, OrLn1, OrLn2, OrRmPtr, OrQty, qOrCarats, qOrWt, qOrPureWt, " + _
                "OrSalRt, OrCstRt, OdPrdQty, qOdBalPrd, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo, OdDmCd, OdDmSz, " + _
                "OdSfx, OmCmCd, OmPoNo, DmCtg, OmDt, OdDelDt, IsNull(RrRngSz, '') as qRrRngSz, '' as qSLoc, " + _
                "0 as qStkWt, 0 as qStkQty /*, 0 as qSRmStkRt */ From #TmpRmReq "

  '**** STK recs shld come first as we need RrRngSz which is obtained from STK recs only
  ws_OrdBy = " order by qTag /*desc*/, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo "
  wSqlStrg = wSqlStrg + ws_OrdBy
  
  With REP
    If UCase$(adc("UoRmCtgFr")) = "D" Or UCase$(adc("UoRmCtgFr")) = "C" Then
      .PHDPureWt.Suppress = True: .PHDRmWt.Suppress = True
      '.PHGPureWt.Suppress = True: .PHGRmWt.Suppress = True
      .DPureWt.Suppress = True: .DRmWt.Suppress = True
      .G3FPureWt.Suppress = True: .G3FRmWt.Suppress = True
      .G2FPureWt.Suppress = True: .G2FRmWt.Suppress = True
      .G1FPureWt.Suppress = True: .G1FRmWt.Suppress = True
      .RFPureWt.Suppress = True: .RFRmWt.Suppress = True
    Else
      .PHDPcRmQty.Suppress = True: .PHDPcCarats.Suppress = True
      '.PHGPcRmQty.Suppress = True: .PHGPcCarats.Suppress = True
      .DPcRmQty.Suppress = True: .DPcCarats.Suppress = True
    End If
  End With
  ShowRepSummFld REP.Fldr1BalPrd, REP.Fldr2BalPrd, REP.Fldr3BalPrd, REP.FldrBalPrd, GRP_REP.GrpOpts, "Rm Code,Rm Ctg,Rm SubCtg,Sub Shp,Parent Cd,Rm Len/Sv,Rm Len 2,Rm Ptr,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt"
  ShowRepSummFld REP.G1FPrdQty, REP.G2FPrdQty, REP.G3FPrdQty, REP.RFPrdQty, GRP_REP.GrpOpts, "Rm Code,Rm Ctg,Rm SubCtg,Sub Shp,Parent Cd,Rm Len/Sv,Rm Len 2,Rm Ptr,RmCd Sale Rate,RmCd Cost Rate,(RmCd) SaleRt,(RmCd) CostRt"
  
  adc.RepRecSource = wSqlStrg
'  Debug.Print wSqlStrg
  
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                    "HCd='" + ctSelfCmCd + "'") = "N" Then _
     REP.hOrRmPtr.Suppress = True: REP.FldOrRmPtr.Suppress = True

  If adc("UOYN2") = "N" Then
    REP.BOXORDHD.Suppress = True: REP.BOXORDDET.Suppress = True: REP.BOXORDDETFT.Suppress = True
    REP.BOXFT1.Suppress = True
  End If
  If adc("UOYN4") = "N" Then
    REP.BOXSTKHD.Suppress = True: REP.BOXSTKDET.Suppress = True: REP.BOXSTKDETFT.Suppress = True
    REP.BOXFT2.Suppress = True
  End If
  CRV_REP.DisplayGroupTree = False
  CmdSC.Top = Me.ScaleHeight - CmdSC.Height
End Sub
'''' std code not to be changed
Private Sub ADC_Load()
    Set moCn = adc.Connection
    Set adc.FirNKeyCtl = adc("UoRmCtgFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gb_RmCtgFor = True
    gb_RmZ = False
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "SO"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub


Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
  Select Case UCase(IdName)
  ' urmi 2.06 Pwd for Val
  Case Is = UCase("UoYn")
    If pv_NewValue <> "" And Trim(txtPwd.text) <> moCn.GetFldVal("Select PValue1 from Param where " + _
                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
      Cancel = True: DispMsg "Enter SEO Password to View Values", etError
    End If
  Case Is = UCase("UoYn1"), UCase("UoYn3")
    If pv_NewValue <> "" And Not moCn.RecSeek("Select 1 from Param where " + _
                           "Ptyp = 'YN' and PmCd = '" + pv_NewValue + "'") Then
      Cancel = True: ErrMsg = "Value Not Defined. (Enter 'Y'es/'N'o)"
    End If
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      Cancel = IIF(adc("UoYN3") = "N", True, False)
      ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used"
    Case Is = UCase("UoIssLocFr"), UCase("UoIssLocTo"), UCase("UoIssLocSel")
      gs_LocTyp = "'S'"
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      If adc("UoYN3") = "N" Then adc("UoYn1") = "N"
  End Select
End Sub

Private Sub CmdSC_Click()
  If FraSC.Visible = False Then
    CmdSC.BackColor = vbWhite
    CmdSC.FontBold = True
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    FraSC.ZOrder
    adc("UOYN2").SetFocus
  Else
    CmdSC.BackColor = &H8000000F
    CmdSC.FontBold = False
    TAB_REP.Enabled = True
    FraSC.Enabled = False
    FraSC.Visible = False
  End If
End Sub

'urmi 2.06
Private Sub TxtPwd_Validate(Cancel As Boolean)
 If Trim(txtPwd.text) = "" Then adc("UoYn") = ""
 
    If Trim(txtPwd.text) = "" Or Trim(txtPwd.text) = moCn.GetFldVal("Select PValue1 from Param where " + _
                           "Ptyp = 'SEOPWD' and PmCd = 'SEOPWD'") Then
    
      Exit Sub
    Else
      Cancel = True
      DispMsg "Invalid Password", etError
    End If
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
 adc("UOYN") = ""
 If adc("UORMCTGFR") = "" Then adc("UORMCTGFR") = "D"
 If adc("UODTFR") = "" Then adc("UODTFR") = MWLib.EmptyDate
 If adc("UODTTO") = "" Then adc("UODTTO") = MWLib.EmptyDate
 If adc("UODELDTFR") = "" Then adc("UODELDTFR") = MWLib.EmptyDate
 If adc("UODELDTTO") = "" Then adc("UODELDTTO") = MWLib.EmptyDate
 If adc("UOYN2") = "" Then adc("UOYN2") = "N"
  If adc("UOYN4") = "" Then adc("UOYN4") = "N"
 If adc("UOYYMMFR") = "0" Then adc("UOYYMMFR") = "198001"
 If adc("UOYYMMTO") = "0" Then adc("UOYYMMTO") = "202812"
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                    "HCd = '" + ctSelfCmCd + "' ")
  REP.wCoCd.SetText ws_HName
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
    Set moCn = Nothing
    Set REP = Nothing
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
Private Sub TmpTbl()
  Dim wStr As String
   '*** Creating Temporary Table #TmpTbl
  wStr = "Create Table  #TmpRmReq " + _
          "(RrRngSz VarChar(12), Grp1 VarChar(30)," + _
          "Grp2 VarChar(30)," + _
          "Grp3 VarChar(30)," + _
          "DGrp1 VarChar(30)," + _
          "DGrp2 VarChar(30)," + _
          "DGrp3 VarChar(30)," + _
          "G1Desc VarChar(30)," + _
          "G2Desc VarChar(30)," + _
          "G3Desc VarChar(30)," + _
          " OrRmCd  VarChar(12)," + _
          "OrRmCtg VarChar(2), " + _
          "OrRmSCtg VarChar(5), " + _
          "qPtrYN  VarChar(1), " + _
          "OrLn1  Float, OrLn2 Float , " + _
          "OrRmPtr Float ,OrQty  Integer ," + _
          "qOrCarats Float,qOrWt Float,qOrPureWt Float, " + _
          "OrSalRt Float,OrCstRt Float, OdPrdQty Float," + _
          "qOdBalPrd Float,OrTc VarChar(3),OrYy Varchar(2),OrChr Varchar(3), " + _
          "OrNo Integer,OrSr Integer, OrSrNo Integer, OdDmCd VarChar(15), " + _
          "OdDmSz VarChar(5), OdSfx VarChar(10),OmCmCd  VarChar(8), OmPoNo VarChar(20)," + _
          "DmCtg VarChar(5), OmDt SmallDateTime, OdDelDt SmallDateTime ) "
   moCn.CreateTmpTable "#TmpRmReq", wStr
  
End Sub




