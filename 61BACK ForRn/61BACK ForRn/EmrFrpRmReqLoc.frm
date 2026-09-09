VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlforspeed8.ocx"
Begin VB.Form EmrFrpRmReqLoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location Wise Raw Material Requirement"
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
      TabIndex        =   96
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
      Left            =   11340
      TabIndex        =   95
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
      TabIndex        =   94
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
      Left            =   5325
      TabIndex        =   104
      Top             =   9660
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   105
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
         TabIndex        =   106
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
      Left            =   6540
      TabIndex        =   103
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   100
         Top             =   -15
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
      Height          =   9675
      Left            =   0
      TabIndex        =   99
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   98
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpRmReqLoc.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRmReqLoc.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   102
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   97
               Top             =   360
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   101
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptWHSet 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1485
               TabIndex        =   157
               ToolTipText     =   "Select Wax Set / Hand Set / Both"
               Top             =   8260
               Width           =   5250
               Begin VB.OptionButton OptWHSet 
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
                  Left            =   3240
                  TabIndex        =   90
                  ToolTipText     =   "Select Wax Set / Hand Set / Both"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptWHSet 
                  Caption         =   "Hand Set"
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
                  Left            =   1560
                  TabIndex        =   89
                  ToolTipText     =   "Select Wax Set / Hand Set / Both"
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton OptWHSet 
                  Caption         =   "Wax Set"
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
                  TabIndex        =   88
                  ToolTipText     =   "Select Wax Set / Hand Set / Both"
                  Top             =   0
                  Width           =   1305
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Det "
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
               Left            =   0
               TabIndex        =   78
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   7665
               Width           =   1680
            End
            Begin VB.Frame FraOptPtr 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   11760
               TabIndex        =   151
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   7950
               Width           =   2385
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
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   8790
               TabIndex        =   150
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   7950
               Width           =   1485
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
                  TabIndex        =   84
                  ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
                  Top             =   0
                  Width           =   1245
               End
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
                  TabIndex        =   85
                  ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
                  Top             =   285
                  Width           =   1245
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4140
               TabIndex        =   19
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1845
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
               TabIndex        =   20
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1850
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
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1245
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
               TabIndex        =   14
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1245
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
               Top             =   405
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
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1245
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
               TabIndex        =   18
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1845
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
               TabIndex        =   25
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   2460
               Width           =   1035
               _ExtentX        =   1826
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
               TabIndex        =   24
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   2460
               Width           =   1035
               _ExtentX        =   1826
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
               TabIndex        =   27
               ToolTipText     =   "Enter To Raw Material Breadth"
               Top             =   2745
               Width           =   675
               _ExtentX        =   1191
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
               TabIndex        =   26
               ToolTipText     =   "Enter From Raw Material Breadth"
               Top             =   2745
               Width           =   675
               _ExtentX        =   1191
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
               TabIndex        =   29
               ToolTipText     =   "Enter To Sale Rate"
               Top             =   3030
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
               TabIndex        =   28
               ToolTipText     =   "Enter From Sale Rate"
               Top             =   3030
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
               TabIndex        =   31
               ToolTipText     =   "Enter To Cost Rate"
               Top             =   3315
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
               TabIndex        =   30
               ToolTipText     =   "Enter From Cost Rate"
               Top             =   3315
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
               TabIndex        =   33
               ToolTipText     =   "Enter To Production Quantity"
               Top             =   3600
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
               TabIndex        =   32
               ToolTipText     =   "Enter From Production Quantity"
               Top             =   3600
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
               TabIndex        =   35
               ToolTipText     =   "Enter To Balance Production"
               Top             =   3885
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
               TabIndex        =   34
               ToolTipText     =   "Enter From Balance Production"
               Top             =   3885
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
               TabIndex        =   48
               ToolTipText     =   "Enter To Design Category"
               Top             =   4740
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
               TabIndex        =   49
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   4740
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
               TabIndex        =   47
               ToolTipText     =   "Enter From Design Category"
               Top             =   4740
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
               TabIndex        =   51
               ToolTipText     =   "Enter To Design Code"
               Top             =   5025
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
               TabIndex        =   52
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   5025
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
               TabIndex        =   50
               ToolTipText     =   "Enter From Design Code"
               Top             =   5025
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
               TabIndex        =   55
               ToolTipText     =   "Enter To Design Size"
               Top             =   5310
               Width           =   705
               _ExtentX        =   1244
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
               TabIndex        =   53
               ToolTipText     =   "Enter From Design Size"
               Top             =   5310
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "OdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4920
               TabIndex        =   56
               ToolTipText     =   "Enter To Suffix"
               Top             =   5310
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
               Left            =   2250
               TabIndex        =   54
               ToolTipText     =   "Enter From Suffix"
               Top             =   5310
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
               TabIndex        =   58
               ToolTipText     =   "Enter To Customer"
               Top             =   5595
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
               TabIndex        =   59
               ToolTipText     =   "Enter Customer Selection"
               Top             =   5595
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
               TabIndex        =   57
               ToolTipText     =   "Enter From Customer"
               Top             =   5595
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
               TabIndex        =   61
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   5880
               Width           =   2655
               _ExtentX        =   4683
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
               TabIndex        =   62
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   5880
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
               TabIndex        =   60
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   5880
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
               TabIndex        =   69
               ToolTipText     =   "Enter To Order Date"
               Top             =   6450
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
               TabIndex        =   68
               ToolTipText     =   "Enter From Order Date"
               Top             =   6450
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
               Index           =   5
               Left            =   4905
               TabIndex        =   40
               ToolTipText     =   "Enter To Order Year"
               Top             =   4170
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
               TabIndex        =   37
               ToolTipText     =   "Enter From Order Year"
               Top             =   4170
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
               TabIndex        =   36
               ToolTipText     =   "Enter From Order Tc"
               Top             =   4170
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
               TabIndex        =   38
               ToolTipText     =   "Enter From Order Character"
               Top             =   4170
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
               TabIndex        =   39
               ToolTipText     =   "Enter To Order Tc"
               Top             =   4170
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
               TabIndex        =   41
               ToolTipText     =   "Enter To Order Character"
               Top             =   4170
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
               TabIndex        =   45
               ToolTipText     =   "Enter To Order Serial"
               Top             =   4455
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
               TabIndex        =   43
               ToolTipText     =   "Enter From Order Serial"
               Top             =   4455
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
               TabIndex        =   42
               ToolTipText     =   "Enter From Order Number"
               Top             =   4455
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
               TabIndex        =   44
               ToolTipText     =   "Enter To Order Number"
               Top             =   4455
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
               TabIndex        =   10
               ToolTipText     =   "Enter To Bag Location"
               Top             =   975
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "BLoc<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   6840
               TabIndex        =   11
               ToolTipText     =   "Enter Bag Location Selection"
               Top             =   960
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "BLoc In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1485
               TabIndex        =   9
               ToolTipText     =   "Enter Fr Bag Location"
               Top             =   975
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "BLoc>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   4140
               TabIndex        =   6
               ToolTipText     =   "Enter To Bag Year"
               Top             =   690
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
               Index           =   47
               Left            =   1485
               TabIndex        =   3
               ToolTipText     =   "Enter From Bag Year"
               Top             =   690
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
               Index           =   49
               Left            =   1905
               TabIndex        =   4
               ToolTipText     =   "Enter From Bag Character"
               Top             =   690
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
               Index           =   50
               Left            =   2685
               TabIndex        =   5
               ToolTipText     =   "Enter From Bag Number"
               Top             =   690
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
               Index           =   51
               Left            =   4545
               TabIndex        =   7
               ToolTipText     =   "Enter To Bag Character"
               Top             =   690
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
               Index           =   55
               Left            =   5325
               TabIndex        =   8
               ToolTipText     =   "Enter To Bag Number"
               Top             =   690
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
               Index           =   56
               Left            =   1935
               TabIndex        =   77
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   7665
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   9030
               TabIndex        =   79
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   8070
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOYN1"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   1485
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   405
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4140
               TabIndex        =   71
               ToolTipText     =   "Enter To Prd Delivery Date"
               Top             =   6735
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
               Index           =   60
               Left            =   1485
               TabIndex        =   70
               ToolTipText     =   "Enter From Prd Delivery Date"
               Top             =   6735
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
               Index           =   61
               Left            =   1485
               TabIndex        =   74
               ToolTipText     =   "Enter From  Priority Code"
               Top             =   7350
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
               Index           =   62
               Left            =   4140
               TabIndex        =   75
               ToolTipText     =   "Enter To  Priority Code"
               Top             =   7350
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
               Index           =   63
               Left            =   6840
               TabIndex        =   76
               ToolTipText     =   "Enter  Priority Code Selection"
               Top             =   7350
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
               Index           =   64
               Left            =   13200
               TabIndex        =   80
               ToolTipText     =   "Enter 'Y' For Calculated Pointer, 'N' for Stored Actual Pointer"
               Top             =   8070
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
               Index           =   65
               Left            =   1485
               TabIndex        =   21
               ToolTipText     =   "Enter From Suffix"
               Top             =   2140
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
               Index           =   66
               Left            =   4140
               TabIndex        =   22
               ToolTipText     =   "Enter From Suffix"
               Top             =   2140
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UORNGSZTO"
               IdName          =   "UORNGSZTO"
               CmpStr          =   "RrRngSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   67
               Left            =   6840
               TabIndex        =   23
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2140
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
               Index           =   68
               Left            =   1485
               TabIndex        =   81
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   7950
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
               Index           =   69
               Left            =   2775
               TabIndex        =   64
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   6165
               Width           =   1330
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
               CmpStr          =   "OdPoNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   5430
               TabIndex        =   66
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   6165
               Width           =   1360
               _ExtentX        =   2408
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
               CmpStr          =   "OdPoNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   6840
               TabIndex        =   67
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   6165
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
               Index           =   72
               Left            =   4140
               TabIndex        =   82
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   7950
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
               Index           =   73
               Left            =   5685
               TabIndex        =   83
               ToolTipText     =   "Enter Bag Character"
               Top             =   7950
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTCHRFR"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOTCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   74
               Left            =   4140
               TabIndex        =   2
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   405
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
               Index           =   75
               Left            =   4140
               TabIndex        =   16
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1540
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
               Index           =   76
               Left            =   1485
               TabIndex        =   15
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1540
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
               Index           =   77
               Left            =   6840
               TabIndex        =   17
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1560
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
               Index           =   78
               Left            =   4140
               TabIndex        =   73
               ToolTipText     =   "Enter To Exp Delivery Date"
               Top             =   7050
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
               Index           =   79
               Left            =   1485
               TabIndex        =   72
               ToolTipText     =   "Enter From Exp Delivery Date"
               Top             =   7050
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
               Index           =   80
               Left            =   4140
               TabIndex        =   92
               ToolTipText     =   "Enter To Set Code"
               Top             =   8560
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOLSCDTO"
               IdName          =   "UOLSCDTO"
               CmpStr          =   "OrSetSCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   82
               Left            =   1485
               TabIndex        =   91
               ToolTipText     =   "Enter From Set Code"
               Top             =   8560
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOLSCDFR"
               IdName          =   "UOLSCDFR"
               CmpStr          =   "OrSetSCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   81
               Left            =   6840
               TabIndex        =   93
               ToolTipText     =   "Enter Set Code Selection"
               Top             =   8560
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOLSCDSEL"
               IdName          =   "UOLSCDSEL"
               CmpStr          =   "OrSetSCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   255
               Index           =   83
               Left            =   6840
               TabIndex        =   159
               ToolTipText     =   "Show Setting Type(Wax Set/Hand Set/Both)"
               Top             =   8280
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   450
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   84
               Left            =   1485
               TabIndex        =   63
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   6165
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
               Index           =   85
               Left            =   4140
               TabIndex        =   65
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   6165
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
               Index           =   86
               Left            =   6840
               TabIndex        =   46
               ToolTipText     =   "Enter Order Number Selection"
               Top             =   4455
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOOMNOSEL"
               IdName          =   "UOOMNOSEL"
               CmpStr          =   "OdNo In"
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
               TabIndex        =   161
               Top             =   6165
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
               TabIndex        =   160
               Top             =   6165
               Width           =   105
            End
            Begin VB.Label LblWHSet 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "W / H Set"
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
               Height          =   240
               Left            =   30
               TabIndex        =   158
               Top             =   8235
               Width           =   885
            End
            Begin VB.Label LblSetCd 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Set Cd"
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
               Height          =   240
               Left            =   30
               TabIndex        =   156
               Top             =   8560
               Width           =   630
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
               Index           =   40
               Left            =   30
               TabIndex        =   155
               Top             =   7050
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
               Index           =   39
               Left            =   30
               TabIndex        =   154
               Top             =   1540
               Width           =   1425
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
               Left            =   4770
               TabIndex        =   153
               Top             =   7950
               Width           =   885
            End
            Begin VB.Label LblBagTyp 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   3030
               TabIndex        =   152
               ToolTipText     =   "Location"
               Top             =   7950
               Width           =   900
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
               Index           =   38
               Left            =   30
               TabIndex        =   149
               Top             =   6165
               Width           =   1425
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   175
               Left            =   30
               TabIndex        =   148
               Top             =   7950
               Width           =   855
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
               Index           =   37
               Left            =   4860
               TabIndex        =   147
               Top             =   5340
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
               Index           =   36
               Left            =   2190
               TabIndex        =   146
               Top             =   5310
               Width           =   105
            End
            Begin VB.Label ALBL 
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
               Height          =   285
               Index           =   35
               Left            =   10410
               TabIndex        =   145
               Top             =   7950
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   30
               TabIndex        =   144
               Top             =   7350
               Width           =   1005
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
               Index           =   34
               Left            =   30
               TabIndex        =   143
               Top             =   6735
               Width           =   1425
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
               Left            =   2205
               TabIndex        =   142
               Top             =   390
               Width           =   105
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Qty && Wt"
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
               Left            =   6855
               TabIndex        =   141
               Top             =   7950
               Width           =   1905
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Box Det (Y/N)"
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
               Index           =   48
               Left            =   30
               TabIndex        =   140
               ToolTipText     =   "Location"
               Top             =   7635
               Visible         =   0   'False
               Width           =   1485
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
               Left            =   5265
               TabIndex        =   139
               Top             =   690
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
               Left            =   2625
               TabIndex        =   138
               Top             =   690
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
               Index           =   28
               Left            =   4485
               TabIndex        =   137
               Top             =   690
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
               Index           =   23
               Left            =   1845
               TabIndex        =   136
               Top             =   690
               Width           =   105
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
               Index           =   20
               Left            =   30
               TabIndex        =   135
               Top             =   690
               Width           =   1605
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
               Index           =   12
               Left            =   30
               TabIndex        =   134
               Top             =   975
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
               TabIndex        =   133
               Top             =   4155
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
               TabIndex        =   132
               Top             =   4440
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
               Index           =   24
               Left            =   5265
               TabIndex        =   131
               Top             =   4155
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
               Index           =   25
               Left            =   2625
               TabIndex        =   130
               Top             =   4155
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
               Index           =   26
               Left            =   4845
               TabIndex        =   129
               Top             =   4155
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
               Index           =   27
               Left            =   2205
               TabIndex        =   128
               Top             =   4155
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
               Index           =   30
               Left            =   4965
               TabIndex        =   127
               Top             =   4440
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
               Index           =   31
               Left            =   2325
               TabIndex        =   126
               Top             =   4440
               Width           =   105
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
               Left            =   30
               TabIndex        =   125
               Top             =   6450
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
               Left            =   30
               TabIndex        =   124
               Top             =   5880
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
               Left            =   30
               TabIndex        =   123
               Top             =   5595
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
               Left            =   30
               TabIndex        =   122
               Top             =   2140
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
               Left            =   30
               TabIndex        =   121
               Top             =   5295
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
               Left            =   30
               TabIndex        =   120
               Top             =   5010
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
               Left            =   30
               TabIndex        =   119
               Top             =   4725
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
               Left            =   30
               TabIndex        =   118
               Top             =   3870
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
               Left            =   30
               TabIndex        =   117
               Top             =   3585
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
               Left            =   30
               TabIndex        =   116
               Top             =   3300
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
               Left            =   30
               TabIndex        =   115
               Top             =   3015
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
               Left            =   30
               TabIndex        =   114
               Top             =   2730
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
               Left            =   30
               TabIndex        =   113
               Top             =   2445
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
               Left            =   30
               TabIndex        =   112
               ToolTipText     =   "Location"
               Top             =   405
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
               Left            =   30
               TabIndex        =   111
               Top             =   1245
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
               TabIndex        =   110
               Top             =   1850
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
               TabIndex        =   109
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
               Left            =   4140
               TabIndex        =   108
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
               Left            =   1485
               TabIndex        =   107
               Top             =   90
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRmReqLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRmReqLoc
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_WHSet As String
Private Sub SetGroupSort()
  'GRP_REP.Add "Process", "BLoc", "", "BLoc", "hBLoc", "", "", "", "(Select LocDesc from Loc where LocCd= BLoc)"
  GRP_REP.Add "Process", "BLoc", "", "", "", "", "", "", "(Select LocDesc from Loc where LocCoCd = BCoCd and LocCd= BLoc)"
  GRP_REP.Add "Rm Code", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)+Rtrim(OrRmCd)", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hOrRmCtg,hOrRmSCtg,hOrRmCd", "Rm Ctg,Rm SubCtg,Parent Cd", "", "", "SubString(RmDesc, 1, 30) "
  GRP_REP.Add "Rm Ctg", "Rtrim(OrRmCtg)", "", "OrRmCtg", "hOrRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) "
  GRP_REP.Add "Rm SubCtg", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)", "'('+ RTrim(OrRmCtg)+ ') '+ OrRmSCtg", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) "
  GRP_REP.Add "Sub Shp", "OrSubShp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'SUBSHP' and PMCd = OrSubShp) "
  
  GRP_REP.Add "Parent Cd", "Rtrim(OrRmCtg)+Rtrim(OrRmSCtg)+(case when RmParentCd= '' then OrRmCd else RmParentCd end)", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ (case when RmParentCd= '' then OrRmCd else RmParentCd end)", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "
  
  'GRP_REP.Add "Rm Len/Sv", "Str(OrLn1,6,3)", "", "OrLn1", "hOrLn1", "", "", "", ""
  'GRP_REP.Add "Rm Len 2", "str(OrLn1,6,3)+' '+str(OrLn1,6,3)", "'('+ LTrim(str(OrLn1,6,3))+ ') '+ LTrim(str(OrLn2,6,3))", "OrLn1,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
   '****************Geeta**************Emr206*********12/04/04
   Dim wsPtrYN As String
   wsPtrYN = "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)"
   GRP_REP.Add "Rm Sz/Sv", GetRmSzDBFldNm("OrLn1", "OrLn2", "OrLn3"), "", "wLnSv", "hOrLn1", "", "", "", ""
   '***********
  '(Bef 2.12) GRP_REP.Add "Rm Len 2", "(case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else LTrim(Str(OrLn1,6,3)) end)+' '+str(OrLn1,6,3)", "'('+ (case when Round(OrLn1, 3) = 0.01 then '+0' when Round(OrLn1, 3) = 0.02 then '+00' when Round(OrLn1, 3) = 0.03 then '+000' else LTrim(Str(OrLn1,6,3)) end)+ ') '+ LTrim(str(OrLn2,6,3))", "wLnSv,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
  
  '*** (Jen 2.12)
'  GRP_REP.Add "Rm Len 2", "(case when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.01) then '+0' when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.02) then '+00' when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.03) then '+000' else LTrim(Str(OrLn1, 6, 3)) end)+' '+ str(OrLn2, 6, 3)", "'('+ (case when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.01) then '+0' when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.02) then '+00' when convert (decimal (12, 3), OrLn1) = convert (decimal (12, 3), 0.03) then '+000' else LTrim(Str(OrLn1, 6, 3)) end)+ ') '+ LTrim(str(OrLn2, 6, 3))", "wLnSv,OrLn2", "hOrLn1,hOrLn2", "Rm Len/Sv", "", "", ""
  '*** (Jen 2.12)
  
  ' ****** sachin 2.12 ****** '
    GRP_REP.Add "Range Size", "RrRngSz", "", "RrRngSz", "hRrRngSz", "", "", "", ""
  ' ****** sachin 2.12 ****** '
  
  'Manoj 2.10.0
'  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", "Str(OrRmPtr,7,4)", "", "OrRmPtr", "hOrRmPtr", "", "", "", ""
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "'") = "Y" Then _
    GRP_REP.Add "Rm Ptr", Array("Str(OrRmPtr,7,4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End,7,4)"), Array("Str(OrRmPtr, 7, 4)", "Str(Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End, 7, 4)"), "OrRmPtr", "hOrRmPtr", "", "", "", ""
  'Manoj 2.10.0
  
  ' Zubin 213 (Bef 213)
  'GRP_REP.Add "Rm Sale Rate", "Str(OrSalRt, 10, 3)", "", "OrSalRt", "hOrSalRt", "OrCstRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "
  'GRP_REP.Add "Rm Cost Rate", "Str(OrCstRt, 10, 3)", "", "OrCstRt", "hOrCstRt", "OrSalRt", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ '/'+ OrRmCd+ ')' "

  ' Zubin 213 (CurCd considered and Desc disp fld err rectified)
  '****** Sachin 2.13.0 [Q/W Added]
  GRP_REP.Add "Rm Sale Rate", "'['+RmQw+']'+Str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) else OrSalRt end, 10, 3)", "", "OrSalRt", "hOrSalRt", "OrCstRt", "", "", ""
  GRP_REP.Add "Rm Cost Rate", "'['+RmQwCst+']'+Str(case when qBaseCur= 'Y' then (Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) else OrCstRt end, 10, 3)", "", "OrCstRt", "hOrCstRt", "OrSalRt", "", "", ""
  '****** Sachin 2.13.0 [Q/W Added]
    
  'GRP_REP.Add "(RmCd) Len/Sv", "OrRmCd+Str(OrLn1,4,2)", "'('+OrRmCd+') '+LTrim(str(OrLn1,4,2))", "OrRmCd,OrLn1", "hOrRmCd,hOrLn1", "Rm Ctg,Rm SubCtg,Rm Code", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  'GRP_REP.Add "(RmCd) Len 2", "OrRmCd+str(OrLn1,4,2)+' '+str(OrLn1,4,2)", "OrRmCd+' ('+ LTrim(str(OrLn1,4,2))+ ') '+ LTrim(str(OrLn2,4,2))", "OrRmCd,OrLn1,OrLn2", "hOrRmCd,hOrLn1,hOrLn2", "Rm Ctg,Rm SubCtg,Rm Code,(RmCd) Len/Sv", "", "", " '('+ RTrim(OrRmCtg)+ '/'+ OrRmSCtg+ ')' "
  
  GRP_REP.Add "OrderNo", "OdTc+OdYy+OdChr+str(OdNo)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))", "wOrdNo,OmCmCd", "hOrdNo,hOmCmCd", "Order Dt,Customer Code,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
  GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg,(DsgCd)Colour", "", "", "(Select DmDesc From DsgMst Where DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '') "
  GRP_REP.Add "(DsgCd)Colour", "OdDmCd+OdDmCol", "'('+OdDmCd+') '+OdDmCol", "OdDmCd", "hOdDmCd", "Design Ctg,Design Code", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'PRTCD' and PmCd= OdPrtCd) "
  '******Geeta*****************Emr206****06/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.Text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Customer PO", "OmCmCd+OmPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)", "wOrdNo,OmCmCd,OmDt,OmPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo", "Order Dt,Customer Code,OrderNo", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))"
  '*** Jay 3.2.0 [OdPoNo]
  'GRP_REP.Add "Cust Sub PO", "OmCmCd+OmPONo+OdPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' then '/'+ LTrim(OdPONo) else '' end )", "wOrdNo,OmCmCd,OmDt,OmPoNo,OdPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo,hOdPoNo", "Order Dt,Customer Code,OrderNo,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))"
  'vk.19 sub customer added
  GRP_REP.Add "Cust SUB PO", "OmCmCd+OmPONo+OdSubCust+OdPONo", "'('+ OmCmCd+ ') '+ LTrim(OmPONo)+(case when OdPONo<> '' or OdSubCust<>'' then '/'+ LTrim(OdSubCust)+' : '+ LTrim(OdPONo) else '' end )", "wOrdNo,OmCmCd,OmDt,OmPoNo,OdPoNo", "hOrdNo,hOmCmCd,hOmDt,hOmPONo,hOdPoNo", "Order Dt,Customer Code,OrderNo,Customer PO", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ ')'+ OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))"
  '*** Jay 3.2.0 [OdPoNo]
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "(OrdKt)Colour", "OdKt+OdDmCol", "'('+OdKt+') '+OdDmCol", "OdKt,OdDmCol", "", "", "", "", "(Select PDesc from Param where PTyp= 'KT' and PMCd= OdKt)+ ' ('+(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) +')'"
' **** Zubin 211 **** '
  GRP_REP.Add "OrderSr", "OdTc+OdYy+OdChr+str(OdNo)+str(OdSr)", "OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))", "wOrdNo,OmCmCd,OrSr", "hOrdNo,hOmCmCd,hSr", "OrderNo,Order Dt,Customer Code,Customer PO,Design Ctg,Design Code,Priority,(DsgCd)Colour,(OrdKt)Colour", "", "", "'('+convert(varchar(8), OmDt, 5) +'/'+ OmCmCd+ '/'+ OmPONo+ ')' "
' **** Zubin 211 **** '

  'Sachin 4.1.2
  GRP_REP.Add "Exp Del Date", "convert(varchar(10), OdExpDelDt, 112)", "convert(varchar(8), OdExpDelDt, 3)", "", "", "", "", "", ""

  'Amore.1 - Group sort added. In Report, if Wax set then "WS-" as prefix and Hand set then "HS-" as prefix for set code
  If UCase(ADC.MenuCd) = UCase("RepRmReqSet") Then
      GRP_REP.Add "Set Cd", "case when OrWsQty >0 then 'WS-' when OrHsQty >0 then 'HS-' else '' end +OrSetSCd", "", "", "", "", "", "", "case when OrWsQty >0 then 'Wax Set-' when OrHsQty >0 then 'Hand Set-' else '' end +(Select PDesc From Param Where PTyp= 'LABSCD' and PMCd ='SET' and PSCD=OrSetSCd)"
  End If
    
End Sub
Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '****************Geeta**********Emr206*******qPtrYN Included
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'BYy, BChr, BNo, BLoc, OrRmCd, OrRmCtg, OrRmSCtg,PValue3 as qPtrYN, OrLn1, OrLn2,
  'OrRmPtr, OrQty, OrWt as qOrCarats,  OrWt as qOrWt, OrWt as qOrPureWt,
  'OrSalRt, OrCstRt, OdPrdQty, BQty as qBalPrd,
  'OrTc, OrYy,OrChr,OrNo,OrSr, OrSrNo, OdDmCd, OdDmSz, OdSfx,
  'OmCmCd , OmPoNo, DmCtg, OmDt, OdDelDt
  'from Bag, OrdRm, OrdDsg, OrdMst, DsgMst,Param  where 1=2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wPurityWt As String
  Dim ws_RmQty As String, ws_RmWt As String, ws_RmPtr As String
  
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  'grpflds = GRP_REP.GrpFldLst
    
' **** Crm 2.12 - 23-02-06 **** '
  Dim wCrmCnd As String
  wCrmCnd = " And OmSubmitYN <> 'N' "
  ' **** Crm 2.12 - 23-02-06 **** '
  
  ' Zubin 213
  Dim ws_Join As String
  Dim wb_BaseCur As Boolean
  
  Dim wAddlCnd As String  '****** Sachin 3.02.0
  
  If ADC("UoYN3") = "Y" Then
    ws_RmQty = "OrPrdQty as OrQty": ws_RmWt = "OrPrdWt"
  Else
    ws_RmQty = "OrQty as OrQty": ws_RmWt = "OrWt"
  End If
    
  wrepcnd = ADC.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  ' Zubin 213
  wb_BaseCur = IIF(UCase(ADC("UoCurCdFr")) = "", True, False)
  wCnd = wCnd + IIF(wb_BaseCur = False, " and OmCmCurCd= '" + ADC("UoCurCdFr") + "'", "")
  'Amore.1- checking condition for W/H set
  wCnd = wCnd + IIF(ADC("UOYN2") = "Y", " and OrWsQty > 0 ", IIF(ADC("UOYN2") = "N", " and OrHsQty > 0 ", ""))
  ws_Join = " join (Select PMCd as qPMCd, (case when " + IIF(wb_BaseCur, "1", "0") + " = 1 then 'Y' " + _
            "else 'N' end) as qBaseCur from Param where PTyp= 'CURNCY') Cur on qPMCd= OmCmCurCd "
  ' Zubin 213
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
  Next i
  
  'Manoj 2.10.0
  If ADC("UoYN1") = "Y" And ADC("UoYN3") = "Y" Then
    ws_RmPtr = " Case When OrPrdQty = 0 Then 0 Else (Case When OrRmCtg in ('C', 'D') Then OrPrdWt / (Case When OrPrdQty = 0 Then 1 Else OrPrdQty End) Else 0 End) End As OrRmPtr "
    grpflds = GRP_REP.GrpFldLst(1)
  Else
    ws_RmPtr = " OrRmPtr "
    grpflds = GRP_REP.GrpFldLst(0)
  End If
  'Manoj 2.10.0

  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True: Rep.BoxFt.Suppress = True: Rep.G4F.Suppress = True
  
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wPurityWt = " /(case when OrRmCtg in ('G', 'P', 'S', 'L') then " + _
                        "IsNull((Select RmPurityWt from Param, RmMst " + _
                        "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue and PTyp= 'RMCTG' and PMCd= OrRmCtg), 1) " + _
                 "When OrRmCtg= 'X' then " + _
                          "IsNull((Select RmPurityWt from Param, RmMst RmP  " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " RmP.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmP.RmCd= PValue and PTyp= 'RMCTG' and " + _
                          "PMCd= IsNull((Select RmB.RmCtg From RmMst RmB where " + IIF(gs_Partition = ctCurrPrtn, " RmB.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmB.RmCd= RmBaseCd), '')), 1) " + _
              "else 1 end) "
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(ADC("UOCMCDFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(ADC("UOCMCDTO")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(ADC("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(ADC("UOCMCDSEL")) + ")", "")
    
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTFR")) <> "01/01/80", " OdOmDt >= '" + Trim(ADC("UODTFR")) + "'", "")
    wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(ADC("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(ADC("UODTTO")) <> "01/01/80", " OdOmDt <= '" + Trim(ADC("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '****************Geeta**************Emr206*********12/04/04
   'qPtrYN  field included in the query
   
    Dim ws_BagJoin As String
    
    ws_BagJoin = " Join Param wBTyp On wBTyp.PTyp= 'BCHR' and wBTyp.PMCd= BChr and wBTyp.PSCd= '' " + _
                 IIF(ADC("UoBagTyp") <> "", " and wBTyp.PValue= '" + ADC("UoBagTyp") + "' ", " ") + _
                 IIF(ADC("UOTCHRFR") <> "", " and wBTyp.PMCd= '" + ADC("UOTCHRFR") + "' ", " ")
   
   ' ****** sachin 2.12 - Range Size included in the query ****** '
    ' Zubin 213 (ws_Join added)
    '****** Sachin 3.02 - Id fields in Joins
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
    ' ***** Manali 3.5.0 - 03/12/08 - rejected and melted bags excluded from the result
    ' ***** Manali 3.8.0 - BQty>0 added in Where Clause to Avoid lines with no requirement
    ' ***** Manali 3.10.0 - 31/03/12 - Pure Wt For 'X'
    
    Dim ws_RmIss As String
    'Amore.1- code for Rm issued in bag, it is used for RmReq Location , not for setting
    ws_RmIss = " and (case when IsNull((select sum(TdRmWt* (Case when TdFrRmLoc='W' and TdFrRmDc='D' then 1 " + _
                                                              "when TdFrRmLoc='W' and TdFrRmDc='C' then -1 " + _
                                                              "when TdToRmLoc='W' and TdToRmDc='D' then 1 " + _
                                                              "when TdToRmLoc='W' and TdToRmDc='C' then -1 " + _
             " Else 0 end)) from Txnd Where TdBIdNo=BIdNo and TdCoCd=BCoCd and TdByy = Byy And TdBChr = BChr And " + _
             " TdBNo=BNo " + IIF(gs_Partition = ctCurrPrtn, " and TdPrtKey=BPrtKey", "") + " and TdRmCd=OrRmCd), 0) > 0 then 0 else 1 end) = 1 "

    wSqlStrg = " Select " + grpflds + ", " + _
             " BYy, BChr, BNo, BLoc, OrRmCd, RmQw, RmQwCst, OrRmCtg, OrRmSCtg, IsNull(RrRngSz,'') as RrRngSz, pRms.PValue3 as qPtrYN," + _
             " OrLn1, OrLn2, OrLn3, " + ws_RmPtr + ", " + ws_RmQty + ", " + _
             " (Case  when (OrRmCtg = 'D' or OrRmCtg = 'C') then " + ws_RmWt + " else 0.0 end) as qOrCarats, " + _
             " (Case  when (OrRmCtg = 'G' or OrRmCtg = 'P' or OrRmCtg = 'S' or OrRmCtg = 'L' or OrRmCtg = 'M' or OrRmCtg = 'X') then " + ws_RmWt + " else 0.0 end) as qOrWt, " + _
             " ((Case when (OrRmCtg ='G' or OrRmCtg ='P' or OrRmCtg ='S' or OrRmCtg ='L') then RmPurityWt When OrRmCtg='X' Then RmBasePurityWt Else 0.00 End)" + wPurityWt + " )*" + ws_RmWt + " as qOrPureWt, " + _
             IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrSalRt/OmCnvFct else OrSalRt*OmCnvFct end) ", "OrSalRt") + " as OrSalRt, " + _
             IIF(wb_BaseCur, " (Case when OmMulDiv= 'M' then OrCstRt/OmCnvFct else OrCstRt*OmCnvFct end) ", "OrCstRt") + " as OrCstRt, " + _
             " OdPrdQty, BQty as qBalPrd, OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo, OdDmCd, OdDmSz, OdSfx, OmCmCd, OmPoNo, DmCtg, OmDt, OdDelDt " + _
             " From Bag join OrdDsg on BOdIdNo=OdIdNo and BCoCd= OdCoCd and BOdTc= OdTc and BOdYy= OdYy and BOdChr= OdChr and BOdNo= OdNo and BOdSr= OdSr " + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey=OdPrtKey", "") + _
             " join OrdRm on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc=OrTc and OdYy=OrYy and OdChr = OrChr and OdNo=OrNo and OdSr=OrSr Join OrdMst on OmCoCd=OrCoCd and " + _
             " OmTc=OrTc and OmYy=OrYy and OmChr = OrChr and OmNo=OrNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OrPrtKey", "") + _
             " Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp='DM' and DmCd = OdDmCd and DmSz = '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey", "") + _
             " join RmMst on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey", "") + " Join Param pRms On pRms.PTyp = 'RMSCTG' ANd pRms.PMCd = RmCtg And pRms.PSCd = RmSCtg " + ws_Join + ws_BagJoin + _
             " Left Outer Join RmRt On RrTcTyp='RSZ' and RrCmCtg= 'C' And RrCmCd='" + ctSelfCmCd + "' And RrCtg=RmCtg And RrSCtg=RmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' and " + _
             " RrLsCtg= '' and RrLabMCd= '' And RrFrLn<= (Case When  pRms.PValue3='N' Then OrLn1 Else OrRmPtr End) and RrToLn>= (Case When pRms.PValue3='N' Then OrLn1 Else OrRmPtr End) " + _
             " where BCls= 'N' and BQty>0 and  BRjQty = 0 and (IsNull((Select TdTc from Txnd where BLstYyKey= TdYyKey and  TdPrtKey=BPrtKey and TdBIdNo=BIdno and TdCoCd=BCoCd and  TdByy=Byy and TdBChr=Bchr and TdBNo=BNo and TdSrNo=0),'') <> (Select PValue from Param where PTyp='TCTyp' and PMCd='MLT')) " + _
             IIF(UCase(ADC.MenuCd) <> UCase("RepRmReqSet"), ws_RmIss, "") + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and BPrtKey='" + ctCurrPrtn + "' ", "") + _
             IIF(wAddlCnd <> "", " And " + wAddlCnd, "") + wCrmCnd
                          
             '' bef 3.10.0
             '''"(Case when (OrRmCtg ='G' or OrRmCtg ='P' or OrRmCtg ='S' or OrRmCtg ='L') then (Select RmPurityWt from RmMst where RmCd=OrRmCd)" + wPurityWt + "  else 0.0 end) * " + ws_RmWt + " as qOrPureWt, "
             
  ADC.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt

  If ADC("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.G4F.Suppress = True
  End If
  ' Zubin 213
  If wb_BaseCur Then
    Rep.TxtValIn.SetText "Value In " + moCn.GetFldVal("Select HBaseCurCd from Head " + _
                                        "where HCd= '" + ctSelfCmCd + "' and HCoCd= '" + ADC("UoCoCdFr") + "'") + ""
  Else
    Rep.TxtValIn.SetText "Value In " + UCase(ADC("UoCurCdFr"))
  End If
  
  If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                                                    "HCd='" + ctSelfCmCd + "'") = "N" Then _
     Rep.hOrRmPtr.Suppress = True: Rep.FldOrRmPtr.Suppress = True
     
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OrRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     ' Tag = "ShowBoxDet"
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
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
    Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
    Set ADC.FirNKeyCtl = ADC("UoRmCtgFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gb_RmCtgFor = True
    gb_RmZ = False
    gs_DmTcTyp = "DM"
    gs_OmTcTyp = "SO"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    'Amore.1 - Setting type and setting code are visible only RM Requirement(Setting).
    If UCase(ADC.MenuCd) <> UCase("RepRmReqSet") Then
        LblWHSet.Visible = False
        FraOptWHSet.Visible = False
        LblSetCd.Visible = False
        ADC("UoLsCdFr").Visible = False
        ADC("UoLsCdTo").Visible = False
        ADC("UoLsCdSel").Visible = False
    End If
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '578-28 validating order no. selection should be a number
  Select Case UCase(IdName)
  Case Is = UCase("UoOmNoSel")
  Dim wOmNoSel() As String, i As Integer
    wOmNoSel = Split(pv_NewValue, ",")
    If IsArray(wOmNoSel) Then
      For i = 0 To UBound(wOmNoSel)
        cancel = Not IsNumeric(wOmNoSel(i))
        If cancel = True Then ErrMsg = "Value(" + CStr(i + 1) + ") Is Not Numeric. Enter Order Numbers like 1,2,3 ": Exit Sub
      Next i
    End If
  End Select
  
  Call SetProp(Me, IdName, Valid, cancel, ErrMsg)
End Sub
Private Sub atxt_FldWhen(Index As Integer, ByVal IdName As String, cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      cancel = IIF(ADC("UoYN3") = "N", True, False)
      ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used"
      
    'Sachin 3.11.0
      Case Is = UCase("UoTChrFr")
      If ADC.Mode = xNorm Then
        If ADC("UoBagTyp") = "" Then cancel = True: ErrMsg = "First Specify Bag Type.": Exit Sub
        HlpList.PMCd "BCHR", "'" + ADC("UoBagTyp") + "'": Exit Sub
      End If
      
      'Amore.1 - For help, Param PTyp should be "SET"
      Case Is = UCase("UoLSCdFr"), UCase("UoLSCdTo"), UCase("UoLSCdSel")
        gs_LabMCd = "SET"
      
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoYn1")
      If ADC("UoYN3") = "N" Then ADC("UoYn1") = "N"
  End Select
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptPtr_Click(Index As Integer)
Select Case Index
  Case Is = 0
    If OptRmPrd(1).Value = True Then
      OptPtr(0).Value = False: OptPtr(1).Value = True: ADC("UoYN1") = "N"
    Else
      ADC("UoYN1") = "Y"
    End If
  Case Is = 1
    ADC("UoYN1") = "N"
    OptPtr(1).Value = True: ADC("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptPtr_GotFocus(Index As Integer)
  DispMsg FraOptPtr.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
Select Case Index
  Case Is = 0
    ADC("UoYN3") = "Y"
  Case Is = 1
    ADC("UoYN3") = "N"
    OptPtr(1).Value = True: ADC("UoYN1") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub

'Amore.1 - Assign setting type to UsrOpt table
Private Sub OptWHSet_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    ADC("UoYN2") = "Y"
  Case Is = 1
    ADC("UoYN2") = "N"
  Case Is = 2
    ADC("UoYN2") = ""
  End Select
End Sub

Private Sub OptWHSet_GotFocus(Index As Integer)
  DispMsg OptWHSet(Index).ToolTipText, etInfo
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    ADC("UOUSRCD") = ADC.UsrCd
    ADC("UOMNUCD") = ADC.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 ADC("UoCoCdFr") = gs_CoCd
 
 ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
 If ADC("UoYN3") = "Y" Then OptRmPrd(0).Value = True
 If ADC("UoYN3") = "N" Then OptRmPrd(1).Value = True
 If ADC("UoYN1") = "Y" Then OptPtr(0).Value = True
 If ADC("UoYN1") = "N" Then OptPtr(1).Value = True
 ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
 
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If ADC("UoYN") = "" Then ADC("UoYN") = "N"
  
  If ADC("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  
  'Amore.1 - Restore last user's selection for setting type
  If ADC("UoYN2") = "Y" Then
    OptWHSet(0).Value = True
  ElseIf ADC("UoYN2") = "N" Then
    OptWHSet(1).Value = True
  ElseIf ADC("UoYN2") = "" Then
    OptWHSet(2).Value = True
  End If
  
  '***** Sachin 3.11.0 - Bag Type in Requirement Report
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    ADC("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    ADC("UoTChrFr").Visible = False
    ADC("UoBagTyp") = "P": ADC("UoTChrFr") = ""
  End If
  Set wHeadChk = Nothing
  
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
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
    ws_Dt = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If ADC.HotKeys(KeyCode, Shift) Then
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
Private Sub ATXT_Validate(Index As Integer, cancel As Boolean)
    cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set Rep = Nothing
    '*** (09/08/05)
End Sub

Private Sub xxx()
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
