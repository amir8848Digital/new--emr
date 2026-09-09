VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpDsgCrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design Card Printing"
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
      Left            =   11220
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   9570
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5355
      TabIndex        =   67
      Top             =   9570
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   68
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
         TabIndex        =   69
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
      Left            =   6570
      TabIndex        =   66
      Top             =   9585
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   64
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
      TabIndex        =   60
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
      TabIndex        =   63
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   62
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpDsgCrd.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpDsgCrd.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Dsg Param    "
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
               Index           =   9
               Left            =   60
               TabIndex        =   57
               Tag             =   "ShowDsgPrm"
               ToolTipText     =   "Check To Show Design Param"
               Top             =   8100
               Width           =   2160
            End
            Begin VB.Frame FraOptValid 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2040
               TabIndex        =   93
               ToolTipText     =   "Select if to Show Design On Hold - 'Yes', 'No', 'Both'"
               Top             =   5520
               Width           =   6705
               Begin VB.OptionButton OptDmValid 
                  Caption         =   "Yes"
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
                  TabIndex        =   39
                  ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.OptionButton OptDmValid 
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
                  Left            =   4725
                  TabIndex        =   41
                  ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
                  Top             =   0
                  Width           =   1530
               End
               Begin VB.OptionButton OptDmValid 
                  Caption         =   "No"
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
                  Left            =   2430
                  TabIndex        =   40
                  ToolTipText     =   "Select if to Show Valid Designs - 'Yes', 'No', 'Both'"
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All Images"
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
               Index           =   7
               Left            =   60
               TabIndex        =   56
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Pictures"
               Top             =   7830
               Width           =   2160
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Comp            "
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
               Index           =   6
               Left            =   60
               TabIndex        =   52
               Tag             =   "ShowComp"
               ToolTipText     =   "Check To Show Design Components"
               Top             =   7230
               Width           =   2160
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show History           "
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
               Left            =   60
               TabIndex        =   50
               Tag             =   "ShowDsgHist"
               ToolTipText     =   "Check To Show Design History"
               Top             =   6960
               Width           =   2160
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Dsg Ana         "
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
               Index           =   5
               Left            =   60
               TabIndex        =   48
               Tag             =   "ShowDsgAna"
               ToolTipText     =   "Check To Show Design Analysis "
               Top             =   6675
               Width           =   2160
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Picture          "
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
               Left            =   60
               TabIndex        =   43
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Pictures"
               Top             =   5850
               Width           =   2160
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Dsg On Hold            "
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
               TabIndex        =   38
               Tag             =   "DmHold"
               ToolTipText     =   "Check To Show Design on Hold"
               Top             =   5220
               Width           =   2160
            End
            Begin VB.Frame FraOptPtr 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1995
               TabIndex        =   90
               ToolTipText     =   "Select if to show Caulated Pointer Or Actual Stored Pointer"
               Top             =   7500
               Width           =   3945
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
                  TabIndex        =   54
                  ToolTipText     =   "Select if to show Caulated Pointer Or Actual Stored Pointer"
                  Top             =   0
                  Width           =   1965
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
                  Left            =   2430
                  TabIndex        =   55
                  ToolTipText     =   "Select if to show Caulated Pointer Or Actual Stored Pointer"
                  Top             =   0
                  Width           =   2175
               End
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   2010
               TabIndex        =   89
               ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
               Top             =   6150
               Width           =   3855
               Begin VB.OptionButton OptRmPrd 
                  Caption         =   "Sales Rm"
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
                  Left            =   2430
                  TabIndex        =   46
                  ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
                  Top             =   0
                  Width           =   1455
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
                  TabIndex        =   45
                  ToolTipText     =   "Select Production Rm Qty, Wt or Sales Rm Qty, Wt"
                  Top             =   0
                  Width           =   1245
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2025
               TabIndex        =   35
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   4605
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UORMGRP"
               IdName          =   "UORMGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   2025
               TabIndex        =   36
               ToolTipText     =   "Enter Level Of Labour Detail To Be Shown"
               Top             =   4890
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLABGRP"
               IdName          =   "UOLABGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2565
               TabIndex        =   42
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   5865
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
               Index           =   40
               Left            =   4515
               TabIndex        =   22
               ToolTipText     =   "Enter To Customer"
               Top             =   2895
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   7020
               TabIndex        =   23
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2895
               Visible         =   0   'False
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   2025
               TabIndex        =   21
               ToolTipText     =   "Enter From Customer"
               Top             =   2895
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4515
               TabIndex        =   2
               ToolTipText     =   "Enter To Design Category"
               Top             =   900
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
               Left            =   7020
               TabIndex        =   3
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   900
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
               Left            =   2025
               TabIndex        =   1
               ToolTipText     =   "Enter From Design Category"
               Top             =   900
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
               Left            =   4515
               TabIndex        =   11
               ToolTipText     =   "Enter To Design Code"
               Top             =   1755
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "DmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7020
               TabIndex        =   12
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   1755
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "DmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   2025
               TabIndex        =   10
               ToolTipText     =   "Enter From Design Code"
               Top             =   1755
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "DmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4515
               TabIndex        =   5
               ToolTipText     =   "Enter To Design Production Category"
               Top             =   1185
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGTO"
               IdName          =   "UODMPRDCTGTO"
               CmpStr          =   "DmPrdCtg <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7020
               TabIndex        =   6
               ToolTipText     =   "Enter Design Production Category Selection"
               Top             =   1185
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMPRDCTGSEL"
               IdName          =   "UODMPRDCTGSEL"
               CmpStr          =   "DmPrdCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2025
               TabIndex        =   4
               ToolTipText     =   "Enter From Design Production Category"
               Top             =   1185
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMPRDCTGFR"
               IdName          =   "UODMPRDCTGFR"
               CmpStr          =   "DmPrdCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   4515
               TabIndex        =   19
               ToolTipText     =   "Enter To Karat"
               Top             =   2610
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "DmKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   7020
               TabIndex        =   20
               ToolTipText     =   "Enter Karat Selection"
               Top             =   2610
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "DmKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   2025
               TabIndex        =   18
               ToolTipText     =   "Enter From Karat"
               Top             =   2610
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "DmKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4515
               TabIndex        =   14
               ToolTipText     =   "Enter To Design Size"
               Top             =   2040
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "DmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   2025
               TabIndex        =   13
               ToolTipText     =   "Enter From Design Size"
               Top             =   2040
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "DmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   4515
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Colour"
               Top             =   2325
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLTO"
               IdName          =   "UODMCOLTO"
               CmpStr          =   "DmCol <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   7020
               TabIndex        =   17
               ToolTipText     =   "Enter Design Colour Selection"
               Top             =   2325
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCOLSEL"
               IdName          =   "UODMCOLSEL"
               CmpStr          =   "DmCol In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2025
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Colour"
               Top             =   2325
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLFR"
               IdName          =   "UODMCOLFR"
               CmpStr          =   "DmCol >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   4515
               TabIndex        =   25
               ToolTipText     =   "Enter To Set Family Code"
               Top             =   3180
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOSETCDTO"
               IdName          =   "UOSETCDTO"
               CmpStr          =   "DmSetCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   7020
               TabIndex        =   26
               ToolTipText     =   "Enter Set Family Code Selection"
               Top             =   3180
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSETCDSEL"
               IdName          =   "UOSETCDSEL"
               CmpStr          =   "DmSetCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2025
               TabIndex        =   24
               ToolTipText     =   "Enter From Set Family Code"
               Top             =   3180
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOSETCDFR"
               IdName          =   "UOSETCDFR"
               CmpStr          =   "DmSetCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4515
               TabIndex        =   27
               ToolTipText     =   "Enter To Loss Category "
               Top             =   3465
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGTO"
               IdName          =   "UOLSCTGTO"
               CmpStr          =   "DmLsCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   7020
               TabIndex        =   28
               ToolTipText     =   "Enter Loss Category Selection"
               Top             =   3465
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLSCTGSEL"
               IdName          =   "UOLSCTGSEL"
               CmpStr          =   "DmLsCtg In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   4515
               TabIndex        =   8
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   1470
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
               CmpStr          =   "DmSalCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   7020
               TabIndex        =   9
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   1470
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
               CmpStr          =   "DmSalCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   2025
               TabIndex        =   7
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   1470
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
               CmpStr          =   "DmSalCtg >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2565
               TabIndex        =   37
               ToolTipText     =   "Show Designs On Hold? (Yes / No)"
               Top             =   5220
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
               CmpStr          =   "DmHld = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   6045
               TabIndex        =   44
               ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
               Top             =   6150
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOYN4"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   2025
               TabIndex        =   0
               ToolTipText     =   "Enter Design Type"
               Top             =   615
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCD"
               IdName          =   "UOPMCD"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2565
               TabIndex        =   49
               ToolTipText     =   "Show Design History"
               Top             =   6945
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
               Index           =   58
               Left            =   6045
               TabIndex        =   53
               ToolTipText     =   "Enter 'Y' For Calculated Pointer, 'N' for Stored Actual Pointer"
               Top             =   7500
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOYN4"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   2565
               TabIndex        =   47
               ToolTipText     =   "Show Design Analysis (Yes / No)"
               Top             =   6675
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   2565
               TabIndex        =   51
               ToolTipText     =   "Show Component Details (Y/N)"
               Top             =   7230
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
               Index           =   26
               Left            =   2025
               TabIndex        =   29
               ToolTipText     =   "Enter From Design Date"
               Top             =   3750
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "DmDsgDt>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   4515
               TabIndex        =   30
               ToolTipText     =   "Enter To Design Date"
               Top             =   3750
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   2
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "DmDsgDt<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   2565
               TabIndex        =   92
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   7815
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN7"
               IdName          =   "UOYN7"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   105
               Left            =   9000
               TabIndex        =   94
               ToolTipText     =   "Show Valid Designs (Y/N/' ')"
               Top             =   5520
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN8"
               IdName          =   "UOYN8"
               CmpStr          =   "DmValidYN="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   110
               Left            =   2025
               TabIndex        =   31
               ToolTipText     =   "Enter From Diamond Weight"
               Top             =   4035
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOGRSWTFR"
               IdName          =   "UOGRSWTFR"
               CmpStr          =   "IsNull(qDiaWt,0)>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4515
               TabIndex        =   32
               ToolTipText     =   "Enter To Diamond Weight"
               Top             =   4035
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOGRSWTTO"
               IdName          =   "UOGRSWTTO"
               CmpStr          =   "IsNull(qDiaWt,0)<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   128
               Left            =   2025
               TabIndex        =   33
               ToolTipText     =   "Enter From Colour Stone Weight"
               Top             =   4320
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UORMWTFR"
               IdName          =   "UORMWTFR"
               CmpStr          =   "IsNull(qCsWt,0)>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   129
               Left            =   4515
               TabIndex        =   34
               ToolTipText     =   "Enter To Colour Stone Weight"
               Top             =   4320
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#00.00"
               MaxLength       =   6
               DataType        =   2
               DataField       =   "UORMWTTO"
               IdName          =   "UORMWTTO"
               CmpStr          =   "IsNull(qCsWt,0)<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   2025
               TabIndex        =   98
               ToolTipText     =   "Enter From Loss Category "
               Top             =   3465
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOLSCTGFR"
               IdName          =   "UOLSCTGFR"
               CmpStr          =   "DmLsCtg >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2565
               TabIndex        =   100
               ToolTipText     =   "Show Design Param"
               Top             =   8100
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN9"
               IdName          =   "UOYN9"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Loss Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   99
               Top             =   3465
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Diamond Wt."
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   97
               Top             =   4035
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "CS Wt"
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
               Index           =   115
               Left            =   90
               TabIndex        =   96
               Top             =   4320
               Width           =   1230
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Valid"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   60
               TabIndex        =   95
               Top             =   5520
               Width           =   1455
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Date"
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
               Index           =   12
               Left            =   90
               TabIndex        =   91
               Top             =   3750
               Width           =   1410
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
               Height          =   450
               Index           =   32
               Left            =   90
               TabIndex        =   88
               Top             =   7500
               Width           =   1695
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Typ"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   87
               Top             =   615
               Width           =   1425
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
               Height          =   525
               Left            =   90
               TabIndex        =   86
               Top             =   6150
               Width           =   1515
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg On Hold"
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
               Left            =   90
               TabIndex        =   85
               ToolTipText     =   "Location"
               Top             =   5220
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sal Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   84
               Top             =   1470
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Set Family Cd"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   83
               Top             =   3180
               Width           =   1365
            End
            Begin VB.Image Pic 
               Height          =   1815
               Left            =   9600
               Top             =   4650
               Visible         =   0   'False
               Width           =   1815
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
               Left            =   7020
               TabIndex        =   82
               Top             =   120
               Width           =   1815
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
               Left            =   90
               TabIndex        =   81
               Top             =   2895
               Visible         =   0   'False
               Width           =   1365
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
               Left            =   90
               TabIndex        =   80
               Top             =   1755
               Width           =   1365
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
               Left            =   90
               TabIndex        =   79
               Top             =   900
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Prd Ctg"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   78
               Top             =   1185
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Kt"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   77
               Top             =   2610
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Size"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   76
               Top             =   2040
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Dsg Colour"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   75
               Top             =   2325
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
               Left            =   2025
               TabIndex        =   74
               Top             =   120
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
               Left            =   4515
               TabIndex        =   73
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label LblRmGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "RM Det Level"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   72
               ToolTipText     =   "Location"
               Top             =   4605
               Width           =   1605
            End
            Begin VB.Label LblLabGrp 
               BackStyle       =   0  'Transparent
               Caption         =   "Lab Det Level"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   71
               ToolTipText     =   "Location"
               Top             =   4890
               Width           =   1605
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   65
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   61
               Top             =   360
               Width           =   9015
               _ExtentX        =   15901
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpDsgCrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepDsgCrd
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
End Sub
Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)

'*** Report Sql ***
'*** Jay 2.13(CT) ***
' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' fields added
' **** Manali 3.6.0 - 17/09/09 - DmSrcDsgCd added
' **** Manali 3.10.0 - 07/03/12 - DmDefSz added
'Select DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, PDesc as qPrdCtgDesc, DmSalCtg, PDesc as qSalCtgDesc, DmBagPcs, DmCmCd, DmHld, DmHldDesc,
'DmParts, DmPartDesc, DmOldCd, DmCol, DmPrdSeq, DmPrdInst, DmDsgBy, DmDsgDt, DmModMkr, DmLstMdf, DmUom, DmVaCtg, DmKt,
'DmLsCtg, DmSetCd, DmRegnExc, DmWaxWt, DmModRunWt, DmSilModWt, DmCasPcWt, DmSrcDsgCd, DmDefSz
'space(4) as qTag, 0 as qMainMet,
'DrSr, DrRmCtg, DrRmSCtg,
'DrRmCd, DrLn1, DrLn2, DrQty, DrRmPtr, DrWt, DrSetSCd, DrSubShp, DrAlyCd, DrWsQty, DrHsQty, DrWt As qGldWt, DrWt As qPlWt, DrWt As qSlWt,
'DrWt as qPdWt, DrWt As qOthMtWt, DrWt As qDiaWt, DrWt As qCSWt, DrWt As qAccWt, DrWt As qGrsWt, DrQty as qDiaQty, DrQty as qCSQty, DrQty as qAccQty,
'DlSr, DlMCd, DlSCd, PDesc as qLabDesc, space(1) as DlQw, DlQty,
'DctSr, DctCd, DctSz, DctChr, DctParts, DctRmCd, DctDmCol, DctRmWt, DctDc, DctPrdSeq, DctDesc, DctMainPrtYn
'From DsgMst, DsgRm, DsgLab, DsgCT, Param Where 1 = 2
'*** Jay 2.13(CT) ***
  
  ' Zubin 212
  'Select DmTcTyp, DmCd, DmSz, DmDesc, DmCtg, DmPrdCtg, PDesc as qPrdCtgDesc,
  'DmSalCtg, PDesc as qSalCtgDesc, DmBagPcs, DmCmCd,
  'DmHld, DmHldDesc, DmParts, DmPartDesc, DmOldCd, DmCol, DmPrdSeq, DmPrdInst,
  'DmDsgBy, DmDsgDt, DmModMkr, DmLstMdf, DmUom, DmVaCtg, DmKt, DmLsCtg, DmSetCd, DmRegnExc,
  'DmWaxWt, DmModRunWt, DmSilModWt, DmCasPcWt, DmSrcDsgCd,
  'space(4) as qTag, 0 as qMainMet, DrSr, DrRmCtg, DrRmSCtg, DrRmCd, DrLn1, DrLn2, DrQty,
  'DrRmPtr, DrWt, DrSetSCd, DrSubShp, DrAlyCd, DrWsQty, DrHsQty, DrWt as qGldWt,
  'DrWt as qPlWt, DrWt as qOthMtWt,DrWt as qDiaWt, DrWt As qCSWt, DrWt As qAccWt,
  'DrWt As qGrsWt, DrQty as qDiaQty, DrQty as qCSQty, DrQty as qAccQty,
  'DlSr, DlMCd, DlSCd, PDesc as qLabDesc, space(1) as DlQw, DlQty,
  'DpSr , DpPrtCd, DpPrtSz, DpQty, DpMainPrtYN
  'From DsgMst, DsgRm, DsgLab, DsgPrt, Param Where 1 = 2
  '*** Report Sql ***
  
  '***********Geeta*************Emr207*******
  '*** Sub Report Sql ***
   'Select DmTcTyp, DmCd, DmSz, DmHistory From DsgMst where 1=2
  '*** Sub Report Sql ***
  '***********Geeta*************Emr207*******
      
  ' Zubin 212
  '*** Sub Report Sql ***
  'Select DaTcTyp, DaCd, DaSz, DaAnaSr, PDesc225, DaAnaCd, Space(45) as qDesc from DsgAna, Param where 1=2
  ' Zubin 212
      
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String
  Dim wCommStr As String, wRmStr As String, wRmGrp As String, wRmCommStr As String
  Dim wLabStr As String, wLabGrp As String, wLabCommStr As String
  Dim ws_RmQty As String, ws_RmWt As String, ws_RmPtr As String
  Dim wCTCommStr As String
  Dim wDCWtStr As String
  
  Dim wDsgPrmStr As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
    
  If adc("UoYN3") = "Y" Then
    ws_RmQty = "DrPrdQty": ws_RmWt = "DrPrdWt"
  Else
    ws_RmQty = "DrQty": ws_RmWt = "DrWt"
  End If
    
  'Manoj 2.10.0
  If adc("UoYN4") = "Y" And adc("UoYN3") = "Y" Then
    ws_RmPtr = "( Case When Max(DrPrdQty) = 0 Then 0 Else (Case When Max(DrRmCtg) in ('C', 'D') Then Sum(DrPrdWt / (Case When DrPrdQty = 0 Then 1 Else DrPrdQty End)) Else 0 End) End )"
  Else
    ws_RmPtr = " max(DrRmPtr) "
  End If
  'Manoj 2.10.0

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  wDCWtStr = ""
  If UCase(adc("UoGrsWtFr")) > 0 Or UCase(adc("UoGrsWtTo")) > 0 Or UCase(adc("UoRmWtFr")) > 0 Or UCase(adc("UoRmWtTo")) > 0 Then _
    wDCWtStr = " Left Outer Join (Select DrTcTyp as qDrTcTyp,DrCd as qDrCd,DrSz as qDrSz,Sum((Case when DrRmCtg='D' Then " + ws_RmWt + " Else 0 End)) as qDiaWt,   " + _
              " Sum((Case when DrRmCtg='C' Then " + ws_RmWt + " Else 0 End)) as qCsWt,Max(DrPrtKey) as qDrPrtKey From DsgRm Group By DrTcTyp,DrCd,DrSz) Dr On qDrTcTyp=DmTcTyp and qDrCd=DmCd and qDrSz=DmSz   "

  '578-12 condition for customer code in DmAllow table added
'  Dim wCustCdCnd As String, wDmAllowCnd As String
'  wCustCdCnd = IIF(ADC("UoCmCdFr") <> "", " and DmwCmCd >= '" + ADC("UoCmCdFr") + "' ", "")
'  wCustCdCnd = wCustCdCnd + IIF(ADC("UoCmCdTo") <> "", " and DmwCmCd <= '" + ADC("UoCmCdTo") + "' ", "")
'  wCustCdCnd = wCustCdCnd + IIF(ADC("UoCmCdSel") <> "", " and DmwCmCd In(" + ADC("UoCmCdSel") + ") ", "")
'
'  wDmAllowCnd = IIF(wCustCdCnd <> "", " and (Select isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='Y'),isnull((Select top 1 DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  wCustCdCnd + " and DmwYN ='N'), (Select DmwYN from DmAllow " + _
                  " Where DmwDmTyp = DmTcTyp and DmwDmCd= DmCd and DmwDmSz = DmSz and DmwTyp='CUST' " + _
                  " and DmwCmCd='ZSELF' and DmwYN ='Y') ))) = 'Y' ", "")
               
    
'*********cRM*****212
' Zubin 212
' **** Manali 3.6.0 - 17/09/09 - DmSrcDsgCd added
' **** Manali 3.10.0 - 07/03/12 - DmDefSz added
'4.1.3.0
  wCommStr = "Select DmTcTyp, DmCd, DmSz, max(DmDesc) as DmDesc, max(DmCtg) as DmCtg, max(DmPrdCtg) as DmPrdCtg, " + _
             "(Select PDesc from Param where Ptyp= 'PRDCTG' and PMCd= max(DmPrdCtg)) as qPrdCtgDesc, " + _
             "max(DmSalCtg) as DmSalCtg, (Select PDesc from Param where Ptyp= 'SALCTG' and PMCd=max(DmCtg) and PSCD=max(DmSalCtg)) as qSalCtgDesc, " + _
             "max(DmBagPcs) as DmBagPcs, max(DmCmCd) as DmCmCd, max(DmValidYn) as DmValidYn, max(DmHld) as DmHld, max(DmHldDesc) as DmHldDesc, " + _
             "max(DmParts) as DmParts, max(DmPartDesc) as DmPartDesc, max(DmOldCd) as DmOldCd, max(DmCol) as DmCol, " + _
             "max(DmPrdSeq) as DmPrdSeq, max(DmPrdInst) as DmPrdInst, max(DmDsgBy) as DmDsgBy, max(DmDsgDt) as DmDsgDt, " + _
             "max(DmModMkr) as DmModMkr, max(DmLstMdf) as DmLstMdf, max(DmUom) as DmUom, max(DmVaCtg) as DmVaCtg, " + _
             "Max(DmKt) as DmKt, Max(DmLsCtg) as DmLsCtg, Max(DmSetCd) as DmSetCd, Max(DmRegnExc) as DmRegnExc, " + _
             "Max(DmWaxWt) as DmWaxWt, Max(DmModRunWt) as DmModRunWt, Max(DmSilModWt) as DmSilModWt, Max(DmCasPcWt) as DmCasPcWt, max(DmSrcDsgCd) as DmSrcDsgCd, max(DmDefSz) as DmDefSz, max(DmPrfVendCd) as DmPrfVendCd, "
  
  Select Case UCase(adc("UoRmGrp"))
  Case Is = UCase("N")
    wRmStr = " 0 as DrSr, '*' as DrRmCtg, '*' as DrRmSCtg, '*' as DrRmCd, 0 as DrLn1, " + _
             "0 as DrLn2, sum(" + ws_RmQty + ") as DrQty, 0 as DrRmPtr, sum(" + ws_RmWt + ") as DrWt, '*' as DrSetSCd, '*' as DrSubShp, " + _
             "'*' as DrAlyCd, sum(DrWsQty) as DrWsQty, sum(DrHsQty) as DrHsQty, "
    wRmGrp = " group by " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey,", "") + "DmTcTyp, DmCd, DmSz, DrSr "
  Case Is = UCase("Sr")
    wRmStr = " DrSr, max(DrRmCtg) as DrRmCtg, max(DrRmSCtg) as DrRmSCtg, max(DrRmCd) as DrRmCd, max(DrLn1) as DrLn1, max(DrLn2) as DrLn2, " + _
             "sum(" + ws_RmQty + ") as DrQty, " + ws_RmPtr + " as DrRmPtr, sum(" + ws_RmWt + ") as DrWt, max(DrSetSCd) as DrSetSCd, max(DrSubShp) as DrSubShp, max(DrAlyCd) as DrAlyCd, " + _
             "sum(DrWsQty) as DrWsQty, sum(DrHsQty) as DrHsQty, "
    wRmGrp = " group by " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey,", "") + "DmTcTyp, DmCd, DmSz, DrSr "
'  Case Is = UCase("RmCdRt")
'    wRmStr = " 0 as OrSrNo, OrRmCtg, OrRmSCtg, OrRmCd, 0 as OrLn1, 0 as OrLn2, " + _
'             "sum(OrQty), sum(OrWt), OrSalRt, '*' as DrSetSCd, 0 as OrSetSalRt, sum(OrSetSalVal), sum(OrSalVal), "
'    wRmGrp = " group by DmTcTyp, DmCd, DmSz, DrRmCtg, DrRmSCtg, DrRmCd, OrSalRt "
  Case Is = UCase("RmCd")
    wRmStr = " 0 as DrSr, DrRmCtg, DrRmSCtg, DrRmCd, 0 as DrLn1, 0 as DrLn2, " + _
             "sum(" + ws_RmQty + ") as DrQty, 0 as DrRmPtr, sum(" + ws_RmWt + ") as DrWt, '*' as DrSetSCd, '*' as DrSubShp, '*' as DrAlyCd, " + _
             "sum(DrWsQty) as DrWsQty, sum(DrHsQty) as DrHsQty, "
    wRmGrp = " group by " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey,", "") + "DmTcTyp, DmCd, DmSz, DrRmCtg, DrRmSCtg, DrRmCd "
  Case Is = UCase("RmSCtg")
    wRmStr = " 0 as DrSr, DrRmCtg, DrRmSCtg, '*' as DrRmCd, 0 as DrLn1, 0 as DrLn2, " + _
             "sum(" + ws_RmQty + ") as DrQty, 0 as DrRmPtr, sum(" + ws_RmWt + ") as DrWt, '*' as DrSetSCd, '*' as DrSubShp, '*' as DrAlyCd, " + _
             "sum(DrWsQty) as DrWsQty, sum(DrHsQty) as DrHsQty, "
    wRmGrp = " group by " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey,", "") + "DmTcTyp, DmCd, DmSz, DrRmCtg, DrRmSCtg "
  Case Is = UCase("RmCtg")
    wRmStr = " 0 as DrSr, DrRmCtg, '*' as DrRmSCtg, '*' as DrRmCd, 0 as DrLn1, 0 as DrLn2, " + _
             "sum(" + ws_RmQty + ") as DrQty, 0 as DrRmPtr, sum(" + ws_RmWt + ") as DrWt, '*' as DrSetSCd, '*' as DrSubShp, '*' as DrAlyCd, " + _
             "sum(DrWsQty) as DrWsQty, sum(DrHsQty) as DrHsQty, "
    wRmGrp = " group by " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey,", "") + "DmTcTyp, DmCd, DmSz, DrRmCtg "
  End Select
  
  ' ****** Sachin 2.12 12/12/05 - Silver added to the query below
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  wRmStr = wRmStr + "sum(case when DrRmCtg= 'G' then " + ws_RmWt + " else 0 end) as qGldWt, " + _
           "sum(case when DrRmCtg= 'P' then " + ws_RmWt + " else 0 end) as qPlWt, " + _
           "sum(case when DrRmCtg= 'S' then " + ws_RmWt + " else 0 end) as qSlWt, " + _
           "sum(case when DrRmCtg= 'L' then " + ws_RmWt + " else 0 end) as qPdWt, " + _
           "sum(case when DrRmCtg= 'M' then " + ws_RmWt + " else 0 end) as qOthMtWt, " + _
           "sum(case when DrRmCtg= 'D' then " + ws_RmWt + " else 0 end) as qDiaWt, " + _
           "sum(case when DrRmCtg= 'C' then " + ws_RmWt + " else 0 end) as qCSWt, " + _
           "sum(case when DrRmCtg= 'X' then " + ws_RmWt + " else 0 end) as qAccWt, " + _
           "sum(" + ws_RmWt + "/ (case when DrRmCtg= 'D' or DrRmCtg= 'C' then 5 else 1 end)) as qGrsWt, " + _
           "sum(case when DrRmCtg= 'D' then " + ws_RmQty + " else 0 end) as qDiaQty, " + _
           "sum(case when DrRmCtg= 'C' then " + ws_RmQty + " else 0 end) as qCSQty, " + _
           "sum(case when DrRmCtg= 'X' then " + ws_RmQty + " else 0 end) as qAccQty, "
  
  '4.1.3.0
  If adc("UoYn8") = "Y" Then
    wCnd = wCnd + IIF(wCnd <> "", " And ", "") + " DmValidYN='Y' "
  ElseIf adc("UoYn8") = "N" Then
    wCnd = wCnd + IIF(wCnd <> "", " And ", "") + " DmValidYN='N' "
  End If
  
  '*** Jay 2.13(CT) ***
  '****** Sachin 3.02 - Id fields in Joins
  wRmCommStr = "'1RM' as qTag, (case when max(DrMainMet)= 'Y' or max(DrRmCd)= " + _
           "IsNull((Select DrRmCd From DsgRm where " + IIF(gs_Partition = ctCurrPrtn, " DrPrtKey=DmPrtKey and ", "") + " DrTcTyp= DmTcTyp and DrCd= DmCd and " + _
           "DrSz= DmSz and DrMainMet= 'Y'), '') then 1 else 2 end) as qMainMet, " + wRmStr + _
           "0 as DlSr,'' as DlMCd, '' as DlSCd, " + _
           "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= 'SET' and " + _
           "PSCd= max(DrSetSCd)) as qLabDesc, '' as DlQw, 0 as DlQty, " + _
           "0 as DctSr, '' as DctCd, '' as DctSz, '' as DctChr, '' as DctParts, '' as DctRmCd, " + _
           "'' as DctDmCol, '' as DctRmWt, '' as DctDc, '' as DctPrdSeq, '' as DctDesc, '' as DctMainPrtYn " + _
           "From DsgRm " + _
           "join DsgMst on DmIdNo=DrDmIdNo and DmTcTyp= DrTcTyp and DmCd= DrCd and DmSz= DrSz " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=DrPrtKey ", "") + _
           "Join Param On PTyp = 'RMSCTG' ANd PMCd = DrRmCtg And PSCd = DrRmSCtg " + _
           wDCWtStr + " Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + " DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + wRmGrp
  
  'wRmCommStr = "'1RM' as qTag, (case when max(DrMainMet)= 'Y' or max(DrRmCd)= " + _
           "IsNull((Select DrRmCd From DsgRm where DrTcTyp= DmTcTyp and DrCd= DmCd and " + _
           "DrSz= DmSz and DrMainMet= 'Y'), '') then 1 else 2 end) as qMainMet, " + wRmStr + _
           "0 as DlSr,'' as DlMCd, '' as DlSCd, " + _
           "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= 'SET' and " + _
           "PSCd= max(DrSetSCd)) as qLabDesc, '' as DlQw, 0 as DlQty, " + _
           "0 as DpSr, '' as DpPrtCd, '' as DpPrtSz, 0 as DpQty, '' as DpMainPrtYN " + _
           "From DsgRm " + _
           "join DsgMst on DmTcTyp= DrTcTyp and DmCd= DrCd and DmSz= DrSz " + _
           "Join Param On PTyp = 'RMSCTG' ANd PMCd = DrRmCtg And PSCd = DrRmSCtg " + _
           "where DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + wRmGrp
           '(index= DsgMst)
  '*** Jay 2.13(CT) ***
    
  Select Case UCase(adc("UoLabGrp"))
  Case Is = UCase("N")
    wLabStr = "0 as DlSr, '*' as DlMCd, '*' as DlSCd, " + _
              "'*' as qLabDesc, '*' as DlQw , sum(DlQty) as DlQty "
    wLabGrp = " group by DmTcTyp, DmCd, DmSz, DlSr  "
  Case Is = UCase("Sr")
    wLabStr = "DlSr, max(DlMCd) as DlMCd, max(DlSCd) as DlSCd, " + _
              "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= max(DlMCd) and " + _
              "PSCd= max(DlSCd))  as qLabDesc, " + _
              "IsNull((Select PValue from Param where PTyp= 'LABSCD' and PMCd= max(DlMCd) and " + _
              "PSCd= max(DlSCd)), '') as DlQw, sum(DlQty) as DlQty "
              '"(case when sum(DlQty)>0 then 'Q' else 'W' end) as DlQw, sum(DlQty) "
    wLabGrp = " group by DmTcTyp, DmCd, DmSz, DlSr "
  Case Is = UCase("LabMCd")
    wLabStr = "0 as DlSr, DlMCd, '*' as DlSCd, " + _
              "'*' as qLabDesc, '*' as DlQw, sum(DlQty) as DlQty "
    wLabGrp = " group by DmTcTyp, DmCd, DmSz, DlMCd "
  Case Is = UCase("LabSCd")
    wLabStr = "0 as DlSr, max(DlMCd) as DlMCd, DlSCd, " + _
              "(select rtrim(PDesc) from Param where PTyp= 'LABSCD' and PMCd= DlMCd and " + _
              "PSCd= DLSCd) as qLabDesc, " + _
              "IsNull((Select PValue from Param where PTyp= 'LABSCD' and PMCd= max(DlMCd) and " + _
              "PSCd= DlSCd), '') as DlQw, sum(DlQty) as DlQty "
              '"(case when sum(DlQty)>0 then 'Q' else 'W' end) as DlQw, sum(DlQty) "
    wLabGrp = " group by DmTcTyp, DmCd, DmSz, DlMCd, DlSCd "
  End Select
  
  '*** Jay 2.13(CT) ***
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.03 - 23/06/08 - Pd Mod - 'L' added
  wLabCommStr = "'2LAB' as qTag, 0 as qMainMet, 0 as DrSr, '' as DrRmCtg, '' as DrRmSCtg, " + _
            "'' as DrRmCd, 0 as DrLn1, 0 as DrLn2, 0 as DrQty, 0 as DrRmPtr, 0 as DrWt, '' as DrSetSCd, " + _
            "'' as DrSubShp, '' as DrAlyCd, 0 as DrWsQty, 0 as DrHsQty, 0 as qGldWt, " + _
            "0 as qPlWt, 0 as qSlWt, 0 as qPdWt, 0 as qOthMtWt, 0 as qDiaWt, 0 as qCSWt, 0 as qAccWt, 0 as qGrsWt, " + _
            "0 as qDiaQty, 0 as qCSQty, 0 as qAccQty, " + wLabStr + ", " + _
            "0 as DctSr, '' as DctCd, '' as DctSz, '' as DctChr, '' as DctParts, '' as DctRmCd, " + _
            "'' as DctDmCol, '' as DctRmWt, '' as DctDc, '' as DctPrdSeq, '' as DctDesc, '' as DctMainPrtYn " + _
            "From DsgLab " + _
            "join DsgMst on DmIdNo=DlDmIdNo and DmTcTyp= DlTcTyp and DmCd= DlCd and DmSz= DlSz " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=DlPrtKey ", "") + _
            wDCWtStr + " where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' And ", "") + " DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + wLabGrp
  
  'wLabCommStr = "'2LAB' as qTag, 0 as qMainMet, 0 as DrSr, '' as DrRmCtg, '' as DrRmSCtg, " + _
            "'' as DrRmCd, 0 as DrLn1, 0 as DrLn2, 0 as DrQty, 0 as DrRmPtr, 0 as DrWt, '' as DrSetSCd, " + _
            "'' as DrSubShp, '' as DrAlyCd, 0 as DrWsQty, 0 as DrHsQty, 0 as qGldWt, " + _
            "0 as qPlWt, 0 as qSlWt, 0 as qOthMtWt, 0 as qDiaWt, 0 as qCSWt, 0 as qAccWt, 0 as qGrsWt, " + _
            "0 as qDiaQty, 0 as qCSQty, 0 as qAccQty, " + wLabStr + ", " + _
            "0 as DpSr, '' as DpPrtCd, '' as DpPrtSz, 0 as DpQty, '' as DpMainPrtYN " + _
            "From DsgLab " + _
            "join DsgMst on DmTcTyp= DlTcTyp and DmCd= DlCd and DmSz= DlSz " + _
            "where DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + wLabGrp
            '(index= DsgMst)
  
    '*** Jay 2.13(CT) ***
    If (adc("UoYN6")) = "Y" Then
      '****** Sachin 3.02 - Id fields in Joins
      wCTCommStr = "'3CT' as qTag, 0 as qMainMet, 0 as DrSr, '' as DrRmCtg, '' as DrRmSCtg, " + _
            "'' as DrRmCd, 0 as DrLn1, 0 as DrLn2, 0 as DrQty, 0 as DrRmPtr, 0 as DrWt, '' as DrSetSCd, " + _
            "'' as DrSubShp, '' as DrAlyCd, 0 as DrWsQty, 0 as DrHsQty, 0 as qGldWt, " + _
            "0 as qPlWt, 0 as qSlWt, 0 as qPdWt, 0 as qOthMtWt, 0 as qDiaWt, 0 as qCSWt, 0 as qAccWt, 0 as qGrsWt, " + _
            "0 as qDiaQty, 0 as qCSQty, 0 as qAccQty, " + _
            "0 as DlSr,'' as DlMCd, '' as DlSCd, '' as qLabDesc, '' as DlQw, 0 as DlQty, " + _
            "DctSr, max(DctCd)as DctCd, max(DctSz) as DctSz, max(DctChr) as DctChr, Max(DctParts) as DctParts, " + _
            "Max(DctRmCd) as DctRmCd, Max(DctDmCol) as DctDmCol, Max(DctRmWt) as DctRmWt, Max(DctDc) as DctDc, " + _
            "Max(DctPrdSeq) as DctPrdSeq, Max(DctDesc) as DctDesc, max(DctMainPrtYn) as DctMainPrtYn " + _
            "From DsgCT " + _
            "Join DsgMst on DmIdNo=DctDmIdNo and DmTcTyp= DctTcTyp and DmCd= DctDmCd and DmSz= DctDmSz " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=DctPrtKey ", "") + _
            wDCWtStr + " where " + IIF(gs_Partition = ctCurrPrtn, "DmPrtKey='" + ctCurrPrtn + "' And ", "") + " DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + _
            " Group By DmTcTyp, DmCd, DmSz, DctSr "
              
      wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
             wCommStr + wLabCommStr + " Union All " + _
             wCommStr + wCTCommStr + " order by DmTcTyp, DmCd, DmSz, qTag, qMainMet, DrSr, DlSr, DctChr "
                   
      If moCn.RecSeek("Select hCompBagsYn from Head where HCd='" + ctSelfCmCd + "' And " + _
                      " HCoCd='" + ctSelfCoCd + "' and hCompBagsYn= 'Y'") Then
        Rep.COMPHEAD2.Suppress = True
        Rep.Line29.Suppress = True
        Rep.Line30.Suppress = True
      Else
        Rep.COMPHEAD1.Suppress = True
        Rep.Line27.Suppress = True
        Rep.Line22.Suppress = True
      End If
    Else
       wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
               wCommStr + wLabCommStr + " order by DmTcTyp, DmCd, DmSz, qTag, qMainMet, DrSr, DlSr "
      Rep.COMPHEAD1.Suppress = True
      Rep.COMPHEAD2.Suppress = True
      Rep.COMPDET.Suppress = True
    End If
    '*** Jay 2.13(CT) ***

  'wPrtCommStr = "'3PRT' as qTag, 0 as qMainMet, 0 as DrSr, '' as DrRmCtg, '' as DrRmSCtg, " + _
            "'' as DrRmCd, 0 as DrLn1, 0 as DrLn2, 0 as DrQty, 0 as DrRmPtr, 0 as DrWt, '' as DrSetSCd, " + _
            "'' as DrSubShp, '' as DrAlyCd, 0 as DrWsQty, 0 as DrHsQty, 0 as qGldWt, " + _
            "0 as qPlWt, 0 as qSlWt, 0 as qOthMtWt, 0 as qDiaWt, 0 as qCSWt, 0 as qAccWt, 0 as qGrsWt, " + _
            "0 as qDiaQty, 0 as qCSQty, 0 as qAccQty, " + _
            "0 as DlSr,'' as DlMCd, '' as DlSCd, '' as qLabDesc, '' as DlQw, 0 as DlQty, " + _
            "DpSr, Max(DpPrtCd), Max(DpPrtSz), Sum(DpQty), Max(DpMainPrtYN) " + _
            "From DsgPrt " + _
            "Join DsgMst on DmTcTyp= DpTcTyp and DmCd= DpCd and DmSz= DpSz " + _
            "where DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd + _
            " Group By DmTcTyp, DmCd, DmSz, DpSr "
  '*** Jay 2.13(CT) ***
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added

  ' Error Rectified (10-11-06)
  'wrepcnd = adc.RepCond
  'wCnd = IIF(wrepcnd <> "", " Where " + wrepcnd, "")
  
  wSqlStrg = "Select DmTcTyp, DmCd, DmSz, DmHistory From DsgMst " + wDCWtStr + " where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd
  Call moCn.RepRes(Rep.Subreport1.OpenSubreport, wSqlStrg + ctMaxDopOpt) '****** Sachin 3.01 ctMaxDopOpt added
  
  'Zubin 212
  wSqlStrg = " Select DaTcTyp, DaCd, DaSz, DaAnaSr, ANAFLD.PDesc225, DaAnaCd, IsNull(ANACD.PDesc, DaAnaDesc) as qDesc from DsgAna " + _
             "  Join DsgMst On DmIdNo=DaDmIdNo and DmTcTyp= DaTcTyp and DmCd= DaCd and DmSz= DaSz " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=DaPrtKey And DmPrtKey='" + ctCurrPrtn + "' ", "") + _
             "      Left Outer Join Param ANACD On ANACD.PTyp= 'DAANACD' and ANACD.PMCd= DaAnaSr and ANACD.PSCd= DaAnaCd " + _
             "      Left Outer Join Param ANAFLD On ANAFLD.PTyp= 'DAANAFLD' and ANAFLD.PMCd= DaAnaSr and ANAFLD.PSCd= '' " + _
             wDCWtStr + " Where DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd
            
  Call moCn.RepRes(Rep.SubRptDsgAna.OpenSubreport, wSqlStrg + ctMaxDopOpt)  '****** Sachin 3.01 ctMaxDopOpt added
  If UCase(adc("UoYN5")) = "N" Then Rep.SECDSGANA.Suppress = True
  ' Zubin 212

'*********Bhavna added dsgprm
  wSqlStrg = " select ds.Dptyp, ds.DpCd, ds.DpDmCd, dp.DpDesc, Pdesc,'' as DumDmSz from DsgPrm ds " + _
             " join Param on PTyp='DSGPRM' and PMCd = ds.DpTyp " + _
             " join DsgMst on ds.DpDmCd = DmCd " + _
             " join DsgPrm dp on ds.DpTyp=dp.DpTyp and ds.DpCd=dp.DpCd and dp.DpDmCd = '' " + _
             " where ds.DpDmCd <>'' and DmTcTyp= '" + gs_DmTcTyp + "' and DmSz = '' " + wCnd

 '** wSqlStrg = " Select DpTyp, DpCd, DpDmCd, DpDesc, PDesc from DsgPrm " + _
             "  Join DsgMst On DmCd= DpDmCd " + _
             "      Left Outer Join Param On PTyp= 'DSGPRM' and PMCd= DpTyp " + _
             " Where DmTcTyp= '" + gs_DmTcTyp + "' " + wCnd
            
  Call moCn.RepRes(Rep.SubRptDsgPrm.OpenSubreport, wSqlStrg + ctMaxDopOpt)
  If UCase(adc("UoYn9")) = "N" Then Rep.SecDsgPrm.Suppress = True
'*********Bhavna added dsgprm
  If adc("UoRmGrp") = "N" Then
    Rep.RmHead.Suppress = True
    Rep.RmDet.Suppress = True
  End If
  
  If adc("UoLabGrp") = "N" Then
    Rep.LABHEAD.Suppress = True
    Rep.LABDET.Suppress = True
  End If
  
  If adc("UoYN") = "N" Then
    Rep.Pic3D.Suppress = True
    Rep.PicSk.Suppress = True
  Else
    Rep.Pic3D.Suppress = False
    Rep.PicSk.Suppress = False
  End If
  
  'Sachin 3.11.0 - Additional Images in Design Card
  If adc("UoYN7") = "N" Then
    Rep.AddlImagesSec.Suppress = True
  Else
    Rep.AddlImagesSec.Suppress = False
  End If
  
  '********Geeta******Emr207*********
  If adc("UoYN2") = "Y" Then
    Rep.SecDmHist.Suppress = False
  Else
    Rep.SecDmHist.Suppress = True
  End If
   '********Geeta******Emr207*********
   
  '********Bhavna dsgprm********
  If adc("UoYn9") = "Y" Then
    Rep.SecDsgPrm.Suppress = False
  Else
    Rep.SecDsgPrm.Suppress = True
  End If
  '********Bhavna dsgprm********
   
  
  If UCase(adc.MenuCd) = UCase("RepDMCrd") Then
    Rep.TxtHead.SetText ("Design Detail Card")
  ElseIf UCase(adc.MenuCd) = UCase("RepSMCrd") Then
    Rep.TxtHead.SetText ("Design Sketch Detail Card")
  ElseIf UCase(adc.MenuCd) = UCase("RepPMCrd") Then
    Rep.TxtHead.SetText ("Design Part Detail Card")
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.DrRmCd}", 8)
  GRP_REP.SetFormula Rep, "wRmCdFontSzAly", GetRmcdFontSz("{rdo.DrAlyCd}", 8)
  GRP_REP.SetFormula Rep, "wRmCdFontSzCT", GetRmcdFontSz("{rdo.DctRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
'*** (Jen 2.12)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
  
End Sub

' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 5
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  Case Is = 6
    If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
  Case Is = 7
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  Case Is = 9
    If .Value = Checked Then adc("UoYn9") = "Y" Else adc("UoYn9") = "N"   '******Bhavna added dsgprm
  
  End Select
End With
End Sub
' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added

' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added

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
    Set adc.FirNKeyCtl = adc("UoPMCd")
    '*********Bhavna added dsgcat chk for dsg prm y/n
    Dim mb_DsgCat As Boolean
     mb_DsgCat = moCn.RecSeek("Select 'x' from Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "' and HDsgCatYn='Y' ")
      If mb_DsgCat = True Then
        ChkBoxArr(9).Visible = True: ChkBoxArr(9).Enabled = True
      Else
        ChkBoxArr(9).Visible = False: ChkBoxArr(9).Enabled = False
      End If
    '*********Bhavna added dsgcat chk for dsg prm y/n
    
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_CmCtg = "C"
    gs_RmGrp = "DSGCRD"
    gs_LabGrp = "DSGCRD"
    gs_PTyp1 = "DMTCTYP"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    Call SetGroupSort

End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  'If UCase(adc.MenuCd) = UCase("RepDMCrd") Then
  If adc("UoPMCd") = "DM" Then
    Me.Caption = "Design Details Card"
    gs_DmTcTyp = "DM"
  ElseIf adc("UoPMCd") = "SM" Then
    Me.Caption = "Design Sketch Details Card"
    gs_DmTcTyp = "SM"
  ElseIf adc("UoPMCd") = "PM" Then
    Me.Caption = "Design Part Details Card"
    gs_DmTcTyp = "PM"
  End If
  
  If adc("UoPMCd") = "" Then adc("UoPMCd") = "DM"
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  ' Zubin 212
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  ' Zubin 212
  
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  If adc("UoYN3") = "Y" Then OptRmPrd(0).Value = True
  If adc("UoYN3") = "N" Then OptRmPrd(1).Value = True
  If adc("UoYN4") = "Y" Then OptPtr(0).Value = True
  If adc("UoYN4") = "N" Then OptPtr(1).Value = True
  ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  
  ' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  If adc("UoYN6") = "" Then adc("UoYN6") = "N"
  If adc("UoYN7") = "" Then adc("UoYN7") = "N"
  If adc("UoYn9") = "" Then adc("UoYn9") = "N"    '*********Bhavna added dsgprm
 
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYn9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked   '********Bhavna added dsgprm
  ' ***** Manali 3.5.0 - 22/11/08 - Check Boxes Added
  
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_PTyp1 = "DMTCTYP"
  
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCd")
    If pv_NewValue <> "" And _
        Not moCn.RecSeek(" Select 'x' from Param where PTyp= 'DMTCTYP' " + _
                         " and PMCd= '" + pv_NewValue + "' and PSCd= ''") Then Cancel = True: ErrMsg = "Enter Valid Design Type": Exit Sub
    If pv_NewValue = "DM" Then gs_DmTcTyp = "DM"
    If pv_NewValue = "SM" Then gs_DmTcTyp = "SM"
    If pv_NewValue = "PM" Then gs_DmTcTyp = "PM"
  End Select
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp1 = "DMTCTYP"
  Select Case UCase(IdName)
    Case Is = UCase("UoYN4")
      Cancel = IIF(adc("UoYN3") = "N", True, False)
      ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used"
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoYN4")
    'adc("UoYN4") = "N"   ' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
  End Select
End Sub

Private Sub OptDmValid_Click(Index As Integer)
'4.1.3.0
Select Case Index
  Case Is = 0
    adc("UoYn8") = "Y"
  Case Is = 1
    adc("UoYn8") = "N"
  Case Is = 2
    adc("UoYn8") = ""
End Select
End Sub

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptPtr_Click(Index As Integer)
Select Case Index
  Case Is = 0
    If OptRmPrd(1).Value = True Then
      OptPtr(0).Value = False: OptPtr(1).Value = True: adc("UoYN4") = "N"
      DispMsg "Calc Ptr Can be selected only if Prd Rm Qty and Wt Option is selected ", etError
    Else
      adc("UoYN4") = "Y"
    End If
  Case Is = 1
    adc("UoYN4") = "N"
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
    adc("UoYN3") = "Y"
  Case Is = 1
    adc("UoYN3") = "N": adc("UoYN4") = "N": OptPtr(1).Value = True
End Select
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 05/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
  
  '4.1.3.0
  If adc("UoYn8") = "Y" Then
    OptDmValid(0).Value = True
  ElseIf adc("UoYn8") = "N" Then
    OptDmValid(1).Value = True
  Else
    OptDmValid(2).Value = True
  End If
  
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
