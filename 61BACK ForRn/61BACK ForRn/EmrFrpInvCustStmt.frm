VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvCustStmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Customer Statement"
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
      Left            =   11430
      TabIndex        =   36
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5595
      TabIndex        =   45
      Top             =   9600
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
      Left            =   6810
      TabIndex        =   44
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   41
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
         TabPicture(0)   =   "EmrFrpInvCustStmt.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpInvCustStmt.frx":001C
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
               Width           =   9195
               _ExtentX        =   16219
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
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Sort by Ord Cust"
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
               Index           =   10
               Left            =   60
               TabIndex        =   31
               Tag             =   "InBaseCurr"
               ToolTipText     =   "Check To Sort by Order Cust"
               Top             =   6050
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "In Base Currency     "
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
               TabIndex        =   30
               Tag             =   "InBaseCurr"
               ToolTipText     =   "Check To Show Values in Base Currency"
               Top             =   5760
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Dia && Cs Value /Ct   "
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
               TabIndex        =   20
               Tag             =   "ShowDCValPerCts"
               ToolTipText     =   "Check To Show Diamond and Colr Stone Value Per Carat"
               Top             =   4335
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust PO          "
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
               TabIndex        =   22
               Tag             =   "ShowCustPO"
               ToolTipText     =   "Check To Show Customer PO"
               Top             =   4620
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Picture                     "
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
               TabIndex        =   14
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   3480
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Average Wt    "
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
               TabIndex        =   12
               Tag             =   "ShowAvgWt"
               ToolTipText     =   "Check To Show Average Wts"
               Top             =   3195
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show All Value        "
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
               TabIndex        =   16
               Tag             =   "ShowVal"
               ToolTipText     =   "Check To Show Value"
               Top             =   3765
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Labour Value           "
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
               Left            =   60
               TabIndex        =   18
               Tag             =   "ShowLabVal"
               ToolTipText     =   "Check To Show Labour Value"
               Top             =   4050
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust DsgCd    "
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
               TabIndex        =   24
               Tag             =   "ShowCustDsg"
               ToolTipText     =   "Check To Show Customer Design Code"
               Top             =   4905
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Cust RmCd     "
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
               Index           =   8
               Left            =   60
               TabIndex        =   26
               Tag             =   "ShowCustRm"
               ToolTipText     =   "Check To Show Customer Rm Codes"
               Top             =   5190
               Width           =   2190
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Running Sr    "
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
               Index           =   3
               Left            =   60
               TabIndex        =   28
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Running Serial"
               Top             =   5475
               Width           =   2190
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2835
               TabIndex        =   2
               ToolTipText     =   "Enter  Invoice Year"
               Top             =   915
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   2055
               TabIndex        =   1
               ToolTipText     =   "Enter  Invoice Tc"
               Top             =   915
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3255
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   915
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   2055
               TabIndex        =   4
               ToolTipText     =   "Enter  Invoice Number"
               Top             =   1200
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   2475
               TabIndex        =   13
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   3480
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
               Index           =   4
               Left            =   2475
               TabIndex        =   11
               ToolTipText     =   "Show Average Weights (Yes / No)"
               Top             =   3195
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
               Index           =   17
               Left            =   2055
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   630
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2055
               TabIndex        =   5
               ToolTipText     =   "Enter  Export Number"
               Top             =   1485
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   2475
               TabIndex        =   15
               ToolTipText     =   "Show Value (Yes / No)"
               Top             =   3765
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               ReCalcParent    =   "UOYN4"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2475
               TabIndex        =   17
               ToolTipText     =   "Show Labour Value (Yes / No)"
               Top             =   4050
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               ReCalcOn        =   "UOYN2"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   2055
               TabIndex        =   6
               ToolTipText     =   "Enter First Group"
               Top             =   1770
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5730
               TabIndex        =   32
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   2910
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD1"
               IdName          =   "UOLMCD1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   5730
               TabIndex        =   33
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   3195
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD2"
               IdName          =   "UOLMCD2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   5730
               TabIndex        =   34
               ToolTipText     =   "Enter Lab Main Code"
               Top             =   3480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOLMCD3"
               IdName          =   "UOLMCD3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   2475
               TabIndex        =   19
               ToolTipText     =   "Show Diamond/ct Value  (Y/ N)"
               Top             =   4335
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
               Index           =   15
               Left            =   2475
               TabIndex        =   21
               ToolTipText     =   "Show Customer PO Number (Y/ N)"
               Top             =   4620
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
               Index           =   10
               Left            =   2475
               TabIndex        =   27
               ToolTipText     =   "Show Running Serial"
               Top             =   5475
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
               Index           =   7
               Left            =   2475
               TabIndex        =   23
               ToolTipText     =   "Show Customer Dsg Cd (Yes / No)"
               Top             =   4905
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
               Index           =   16
               Left            =   2475
               TabIndex        =   25
               ToolTipText     =   "Show Custome Rm Cd (Yes / No)"
               Top             =   5190
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN8"
               IdName          =   "UOYN8"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2055
               TabIndex        =   9
               ToolTipText     =   "Enter Level Of Raw Material Detail To Be Shown"
               Top             =   2625
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
               Index           =   8
               Left            =   2055
               TabIndex        =   7
               ToolTipText     =   "Enter Second Group"
               Top             =   2055
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP1"
               IdName          =   "UOINGRP1"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   2055
               TabIndex        =   8
               ToolTipText     =   "Enter Detail Sorting To Be Shown In Report"
               Top             =   2340
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINDET"
               IdName          =   "UOINDET"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2475
               TabIndex        =   29
               ToolTipText     =   "Show Values In Base Currency (Yes / No)"
               Top             =   5760
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN9"
               ReCalcOn        =   "UOINVAL"
               IdName          =   "UOYN9"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2055
               TabIndex        =   10
               ToolTipText     =   "Specify Whether To Show Actual, Invoice Or Customs Values"
               Top             =   2910
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL"
               ReCalcParent    =   "UOYN9"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   2475
               TabIndex        =   75
               ToolTipText     =   "Sort by Order Customer (Yes / No)"
               Top             =   6050
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN10"
               IdName          =   "UOYN10"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "In Base Currency"
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
               Index           =   0
               Left            =   90
               TabIndex        =   74
               ToolTipText     =   "Location"
               Top             =   5760
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.Label LblDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Sort"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   73
               Top             =   2340
               Width           =   1845
            End
            Begin VB.Label LblRunSr 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Running Sr"
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
               Left            =   90
               TabIndex        =   72
               Top             =   5475
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust DsgCd"
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
               Left            =   90
               TabIndex        =   71
               ToolTipText     =   "Location"
               Top             =   4905
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label lblCustRm 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust RmCd"
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
               Left            =   90
               TabIndex        =   70
               ToolTipText     =   "Location"
               Top             =   5190
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label lblRem 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Cust PO "
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
               Left            =   90
               TabIndex        =   69
               Top             =   4620
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Details"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4710
               TabIndex        =   68
               Top             =   2625
               Width           =   1845
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Dia && Cs Value /Ct     "
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
               Left            =   90
               TabIndex        =   67
               Top             =   4335
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               TabIndex        =   66
               Top             =   3480
               Width           =   1845
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               TabIndex        =   65
               Top             =   3195
               Width           =   1845
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               TabIndex        =   64
               Top             =   2910
               Width           =   1845
            End
            Begin VB.Label LblGrpCombo 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination1"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   63
               ToolTipText     =   "Location"
               Top             =   1770
               Width           =   1875
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Grp Combination2"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   62
               ToolTipText     =   "Location"
               Top             =   2055
               Width           =   1875
            End
            Begin VB.Label LblShwVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show All Value"
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
               Left            =   90
               TabIndex        =   61
               ToolTipText     =   "Location"
               Top             =   3765
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label LblLabVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Labour Value"
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
               Left            =   90
               TabIndex        =   60
               ToolTipText     =   "Location"
               Top             =   4050
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Exp No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   59
               Top             =   1485
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
               Index           =   15
               Left            =   90
               TabIndex        =   58
               Top             =   635
               Width           =   1335
            End
            Begin VB.Label LblAvgWt 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Avg Wt"
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
               Left            =   90
               TabIndex        =   57
               Top             =   3195
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Values As"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   56
               Top             =   2910
               Width           =   1845
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Picture"
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
               Left            =   90
               TabIndex        =   55
               Top             =   3480
               Visible         =   0   'False
               Width           =   1845
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
               TabIndex        =   54
               Top             =   2625
               Width           =   1845
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
               Index           =   9
               Left            =   90
               TabIndex        =   53
               Top             =   915
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
               Index           =   10
               Left            =   90
               TabIndex        =   52
               Top             =   1200
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
               Index           =   25
               Left            =   3195
               TabIndex        =   51
               Top             =   915
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
               Left            =   2775
               TabIndex        =   50
               Top             =   915
               Width           =   105
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
               Left            =   4305
               TabIndex        =   49
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
               Left            =   1965
               TabIndex        =   48
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvCustStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepinvCustStmt
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  ' GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
 '**************************** Geeta ***********************
' Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
' *** Manali 3.03 - OdPoNo added, join with OrdDsg added
' **** Manali 3.5.0 - IrRmSz, RisRmSzDesc, qIrLmeRt added
'   Select InCoCd, InYy, InTc, InChr, InNo, InCmCd, CmName as InvCmName,
'   InExpNo, IdExpOdTc, IdExpOdYy,
'   IdExpOdChr, IdExpOdNo, IdExpOdSr, OmCmCd, CmName as OrdCmName, IdSr,
'   IdDmCd, Space(17) as qOurDmCd, DmCtg, DmSalCtg, IdSfx, IdDmSz,
'   Space(5) as OdDmCol, OmPONo, space(15) as CdCustDmCd,
'   IdKt, DmUom, IdGldLs, IdVaCtg, Space(254) as qRmCtgDesc,
'   space(254) as qDmCtgDesc, space(254) as qDmSalCtgDesc, space(254) as qDmColDesc,
'   space(254) as qVaCtgDesc, RmCtg as OrRmCtg, IdIVal as qUnitPrc,
'   IdIVal as IdVal, IdQty, Space(4) as qTag, 0 as qMainMet, IrSrNo,
'   RmCtg, RmSCtg, Space(14) as IrRmCd, RmQw, IrRmSz, RisRmSzDesc, IrRmQty,
'   IrRmIWt, IrRmIRt, qIrLmeRt, IrSetSCd, IrSetIRt, IrSetIVal,
'   IrRmIVal, IlIVal as qTotLabVal, IlIVal as qlab1,
'   IlIVal As qLab2, IlIVal As qlab3
'   From InvHd, InvDsg, InvRm, InvLab, RmMst, DsgMst, OrdMst, CustMst, OrdDsg
'   Where 1 = 2
' *******************************************************
  Dim wrepcnd As String, wCnd As String, wGldAsWt As String, wIdFOB As String, wIrRmWt As String
  Dim wIrRmRt As String, wIrSetRt As String, wIrSetVal As String, wIrRmVal As String
  Dim wiLRt As String, wiLVal As String, wCommStr As String, wRmStr As String, wRmGrp As String
  Dim wCommLabStr As String, wCommRmStr As String
  Dim wRmCommStr As String, wLabStr As String, wLabGrp As String, wLabCommStr As String
  Dim wSqlStrg As String, wDsgCtgStr As String, wRmDet As String
  Dim wRmDetSql As String
  Dim wRtStr As String 'urmi 2.04-3 to get the rate when grouping is by RMCd (TotVal/Wt)
  Dim wOurDsgCd As String, wOurRmCd As String
  
  Dim wIrLmeRt As String, wRmSzDesc As String      ' ***** Manali 3.5.0 - 11/11/08 - qIrLmeRt, RmIdSzDesc added in report [wLnSv]

  Set Rep = Nothing
  Set adc.RepSource = Rep
  
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")

  ' Zubin 213
  Dim ws_BaseCurCd As String, wb_InBaseCur As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  If UCase(adc("UoYN9")) = "Y" Then
    wb_InBaseCur = True
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_InBaseCur = False
    Rep.TxtValIn.SetText "Value In " + _
          moCn.GetFldVal("Select CmCurCd from InvHd " + _
                      "Join CustMst On CmCtg= 'C' and CmCd= InCmCd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and " + _
                      "   ( ( InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
                      "   InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + ") or " + _
                      "         (InExpNo= '" + adc("UOINEXPNOFR") + "'))")
  End If
  ' Zubin 213
  
  If (adc("UoYN2")) = "Y" Then
   If UCase(adc("UoInVal")) = "A" Then
      wGldAsWt = "IdAGldAsWt"
      wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdAVal/InCnvFct else IdAVal*InCnvFct end) ", "IdAVal")
      wIrRmWt = "IrRmAWt"
      wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmARt/InCnvFct else IrRmARt*InCnvFct end) ", "IrRmARt")
      wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) ", "IrSetARt")
      ' Zubin 212 (Before 31/12/05)
      'wIrSetVal = "IrSetAVal"
      wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmAVal/InCnvFct else IrRmAVal*InCnvFct end) ", "IrRmAVal")
      ' Zubin 212 (Before 31/12/05)
      'wiLRt = "iLARt"
      'wiLVal = "iLAVal"
      ' Zubin 212
      If UCase(adc("UoYN4")) = "Y" Then
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
      Else
        wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
      End If
      ' Zubin 212
      ' ***** Manali 3.5.0 - 11/11/08 - IrLmeRt added [wLnSv]
      wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then (case IrRmCtg when 'G' Then OmLmgCst when 'P' then OmLmpCst when 'S' then OmLmsCst when 'L' then OmLmlCst Else 0 End)  " + _
                 " When 'N' then (case IrRmCtg when 'G' Then InLmgCst when 'P' then InLmpCst when 'S' then InLmsCst when 'L' then InLmlCst Else 0 End) " + _
                 " Else 0 end) "

     Rep.TxtInvValues.SetText "(Actual Values)"
   ElseIf UCase(adc("UoInVal")) = "S" Then
      wGldAsWt = "IdAGldAsWt"
      wIdFOB = "0"
      wIrRmWt = "IrRmAWt"
      wIrRmRt = "IrStkRt"
      wIrSetRt = " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) "
      ' Zubin 212 (Before 31/12/05)
      'wIrSetVal = "IrSetAVal"
      wIrRmVal = "IrRmAWt * IrStkRt"
      ' Zubin 212 (Before 31/12/05)
      'wiLRt = "iLARt"
      'wiLVal = "iLAVal"
      ' Zubin 212
      If UCase(adc("UoYN4")) = "Y" Then
        wiLRt = " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) "
        wiLVal = " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) "
        wIrSetVal = " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) "
      Else
        wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
      End If
      ' Zubin 212
      
      wIrLmeRt = "0.00 "    ' ***** Manali 3.5.0 - 11/11/08 - wIrLmeRt added [wLnSv]
     Rep.TxtInvValues.SetText "(Stock Values)"
      
'      If UCase(adc.MenuCd) <> UCase("RepInvCat") And UCase(adc.MenuCd) <> UCase("RepInvCTBCat") Then
'        GRP_REP.SetFormula mRep, "wOdSalPrc", "if {rdo.qOrdQty} > 0 then " + _
'                    "Sum({rdo.OrSalVal}, {rdo.OdSr})+{@g2_labVal}/{rdo.qOrdQty} else 0"
'        GRP_REP.SetFormula mRep, "wSalVal", "Sum({rdo.OrSalVal}, {rdo.OdSr})+{@g2_labVal}"
'      End If
   '**************
   ElseIf UCase(adc("UoInVal")) = "I" Then
      wGldAsWt = "IdiGldAsWt"
      wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdIFob/InCnvFct else IdIFob*InCnvFct end) ", "IdIFob")
      wIrRmWt = "IrRmIWt"
      wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIRt/InCnvFct else IrRmIRt*InCnvFct end) ", "IrRmIRt")
      wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIRt/InCnvFct else IrSetIRt*InCnvFct end) ", "IrSetIRt")
      ' Zubin 212 (Before 31/12/05)
      'wIrSetVal = "IrSetIVal"
      wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIVal/InCnvFct else IrRmIVal*InCnvFct end) ", "IrRmIVal")
      ' Zubin 212 (Before 31/12/05)
      'wiLRt = "iLiRt"
      'wiLVal = "iLiVal"
      ' Zubin 212
      If UCase(adc("UoYN4")) = "Y" Then
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
      Else
        wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
      End If
      ' Zubin 212
      
      ' ***** Manali 3.5.0 - 11/11/08 - qIrLmeRt added Below [wLnSv]
      wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                 " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                 " Else 0 end) "

      Rep.TxtInvValues.SetText "(Invoice Values)"
   ElseIf UCase(adc("UoInVal")) = "Z" Then
      wGldAsWt = "IdZGldAsWt"
      wIdFOB = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IdZFob/InCnvFct else IdZFob*InCnvFct end) ", "IdZFob")
      wIrRmWt = "IrRmZWt"
      wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
      wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZRt/InCnvFct else IrSetZRt*InCnvFct end) ", "IrSetZRt")
      ' Zubin 212 (Before 31/12/05)
      'wIrSetVal = "IrSetZVal"
      wIrRmVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZVal/InCnvFct else IrRmZVal*InCnvFct end) ", "IrRmZVal")
      ' Zubin 212 (Before 31/12/05)
      'wiLRt = "iLZRt"
      'wiLVal = "iLZVal"
      ' Zubin 212
      If UCase(adc("UoYN4")) = "Y" Then
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
      Else
        wiLRt = "0": wiLVal = "0": wIrSetVal = "0"
      End If
      ' Zubin 212
      
      ' ***** Manali 3.5.0 - 11/11/08 - qIrLmeRt added Below [wLnSv]
      wIrLmeRt = " (case  OdIGldRtEqOrd  when 'Y' then OrLmeSal  " + _
                 " When 'N' then (case IrRmCtg when 'G' Then InLmgSal when 'P' then InLmpSal when 'S' then InLmsSal when 'L' then InLmlSal Else 0 End) " + _
                 " Else 0 end) "

      Rep.TxtInvValues.SetText "(Custom Values)"
   End If
  ElseIf (adc("UoYN2")) = "N" Then
      wIdFOB = "0"
      wIrRmVal = "0"
     Rep.TxtInvValues.SetText ""
      wIrLmeRt = "0.00 "      ' ***** Manali 3.5.0 - 11/11/08 - qIrLmeRt added Below [wLnSv]

      If UCase(adc("UoInVal")) = "A" Then
        wGldAsWt = "IdAGldAsWt"
        wIrRmWt = "IrRmAWt"
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmARt/InCnvFct else IrRmARt*InCnvFct end) ", "IrRmARt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) ", "IrSetARt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
     Rep.TxtInvValues.SetText "(Actual Values)"
      
      
      ElseIf UCase(adc("UoInVal")) = "S" Then
        wGldAsWt = "IdAGldAsWt"
        wIrRmWt = "IrRmAWt"
        wIrRmRt = "IrStkRt"
        wIrSetRt = " (Case when InMulDiv= 'M' then IrSetARt/InCnvFct else IrSetARt*InCnvFct end) "
        wiLRt = " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) "
     Rep.TxtInvValues.SetText "(Stock Values)"
    '*****************************
      ElseIf UCase(adc("UoInVal")) = "I" Then
        wGldAsWt = "IdiGldAsWt"
        wIrRmWt = "IrRmIWt"
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmIRt/InCnvFct else IrRmIRt*InCnvFct end) ", "IrRmIRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIRt/InCnvFct else IrSetIRt*InCnvFct end) ", "IrSetIRt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
        Rep.TxtInvValues.SetText "(Invoice Values)"
      
      ElseIf UCase(adc("UoInVal")) = "Z" Then
        wGldAsWt = "IdZGldAsWt"
        wIrRmWt = "IrRmZWt"
        wIrRmRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrRmZRt/InCnvFct else IrRmZRt*InCnvFct end) ", "IrRmZRt")
        wIrSetRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZRt/InCnvFct else IrSetZRt*InCnvFct end) ", "IrSetZRt")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
        Rep.TxtInvValues.SetText "(Custom Values)"
      End If
    
    If (adc("UoYN4")) = "Y" Then
      If UCase(adc("UoInVal")) = "A" Or UCase(adc("UoInVal")) = "S" Then
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetAVal/InCnvFct else IrSetAVal*InCnvFct end) ", "IrSetAVal")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLAVal/InCnvFct else iLAVal*InCnvFct end) ", "iLAVal")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLARt/InCnvFct else iLARt*InCnvFct end) ", "iLARt")
      ElseIf UCase(adc("UoInVal")) = "I" Then
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetIVal/InCnvFct else IrSetIVal*InCnvFct end) ", "IrSetIVal")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiVal/InCnvFct else iLiVal*InCnvFct end) ", "iLiVal")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLiRt/InCnvFct else iLiRt*InCnvFct end) ", "iLiRt")
      ElseIf UCase(adc("UoInVal")) = "Z" Then
        wIrSetVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then IrSetZVal/InCnvFct else IrSetZVal*InCnvFct end) ", "IrSetZVal")
        wiLVal = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZVal/InCnvFct else iLZVal*InCnvFct end) ", "iLZVal")
        wiLRt = IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then iLZRt/InCnvFct else iLZRt*InCnvFct end) ", "iLZRt")
      End If
    Else
        wIrSetVal = "0"
        wiLVal = "0"
        wiLRt = "0"
    End If
  End If

  If UCase(adc("UoYN7")) = "Y" Then
      wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                        "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And ", "") + " CdCtg = 'C' and CdCd = Max(InCmCd) and " + _
                        "CdOurDmCd = Max(IdDmCd) and CdOurSfx = Max(IdSfx) and " + _
                        "CdOurDmSz = Max(IdDmSz)), " + _
                      "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                            "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And ", "") + " CdCtg = 'C' and CdCd = Max(InCmCd) and " + _
                            "CdOurDmCd = Max(IdDmCd) and CdOurSfx = '' and " + _
                            "CdOurDmSz = Max(IdDmSz)), " + _
                          "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And ", "") + " CdCtg = 'C' and CdCd = Max(InCmCd) and " + _
                                "CdOurDmCd = Max(IdDmCd) and CdOurSfx = Max(IdSfx) and " + _
                                "CdOurDmSz = '' ), " + _
                              "IsNull((Select cast(CdCustDmCd as Varchar(40)) from CustDsg " + _
                                    "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' And ", "") + " CdCtg = 'C' and CdCd = Max(InCmCd) and " + _
                                    "CdOurDmCd = Max(IdDmCd) and CdOurSfx = '' and " + _
                                    "CdOurDmSz = ''), '{' + max(IdDmCd) + '}')))) as qOurDmCd"
  Else
    wOurDsgCd = "Max(IdDmCd) as qOurDmCd"
  End If

  If UCase(adc("UoYN8")) = "Y" Then
'        wOurRmCd = "IsNull((Select cast(CrCustRmCd as Varchar(14)) from CustRm where " + IIF(gs_Partition = ctCurrPrtn, " CrPrtKey='" + ctCurrPrtn + "' And ", "") + " CrCtg = 'C' and " + _
'                "CrCd = Max(InCmCd) and CrRmCtg = Max(RmCtg) and " + _
'                "CrOurRmCd = Max(IrRmCd) and 0 Between CrFrLn and CrToLn)," + _
'                "'{' + Max(IrRmCd) + '}') as IrRmCd, Max(RmQw) As RmQw"
        wOurRmCd = " (Case when Max(IrCustRmCd) ='' Then '{'+Max(IrRmCd)+'}' Else Max(IrCustRmCd) End)  as IrRmCd,Max(RmQW) as RmQW  "
  Else
    wOurRmCd = "Max(IrRmCd) as IrRmCd, Max(RmQw) As RmQw"
  End If


  '*********************Assigning Group Desc & group formula ************************************************************
  Dim wStr1 As String, wStr2 As String
  Dim wStr3 As String, wStr4 As String, wStr5 As String, wSqlStrg1 As String, wSqlStrg2 As String
  
  '*** For the Group Formula of the report
  Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
  Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
  Dim i As Integer, qDmCtgDesc As String, qRmCtgDesc As String, qDmColDesc As String, qDmKtDesc As String
  Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String

' **** Zubin 211 **** '
  Dim wDetArr() As String, wDetSrt As String
' **** Zubin 211 **** '
  ' Zubin 212
  Dim qDmSalCtgDesc As String

  qDmCtgDesc = "'' as qDmCtgDesc": qRmCtgDesc = "'' as qRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc": qDmKtDesc = "'' as qDmKtDesc"
  qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
  ' Zubin 212
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
  
  If adc("UoYn10") = "Y" Then wGrpFormula1 = "{rdo.OmCmCd}"     ' 4.1.2
  
  wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp") + "' "), ",")
  For i = 0 To UBound(wGrpArr1)
    If UCase(wGrpArr1(i)) = "RMCTG" Then
      qRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                   "(select RmCtg from RmMst join InvRm on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd=IrRmCd  and  IrMainMet='Y'" + _
                   "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrCocd = InCocd And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                   "and IrNo=InNo and IrSr=IdSr)) as qRmCtgDesc"
      
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
    ElseIf UCase(wGrpArr1(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr1(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr1(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    ElseIf UCase(wGrpArr1(i)) = "PCS/PRS" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr1(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      '*****Gold Loss Description Not Reqd**10/09/04****
'    ElseIf UCase(wGrpArr1(i)) = "GLDLS" Then
'       wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
    ' Zubin 212
    ElseIf UCase(wGrpArr1(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End If
        wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "})))"
    
    Next i
  
  wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp1") + "' "), ",")
  For i = 0 To UBound(wGrpArr2)
    If UCase(wGrpArr2(i)) = "RMCTG" Then
      qRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                   "(select RmCtg from RmMst join InvRm on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd=IrRmCd  and  IrMainMet='Y'" + _
                   "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrCocd = InCocd And IrTc = InTc  and IrYy = InYy and IrChr=InChr " + _
                   "and IrNo=InNo and IrSr=IdSr)) as qRmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
    ElseIf UCase(wGrpArr2(i)) = "KT" Then
      qDmKtDesc = "(Select PDesc from Param where PTyp= 'KT' and PMCd= max(IdKt)) as qDmKtDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmKtDesc})"
    ElseIf UCase(wGrpArr2(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
    ElseIf UCase(wGrpArr2(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    ElseIf UCase(wGrpArr2(i)) = "PCS/PRS" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr2(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    '*****Gold Loss Description Not Reqd**10/09/04****
'    ElseIf UCase(wGrpArr2(i)) = "GLDLS" Then
'      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
    ' Zubin 212
    ElseIf UCase(wGrpArr2(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    ' Zubin 212
    End If
    wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "})))"
  Next i
  '****************************************************************************************

' **** Zubin 211 **** '
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    ' **** Manali 3.6.0 - 12/11/09 - IdTrayNo added
    wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IdTrayNo}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
    ' bef 3.6.0
    'wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
  Next i
' **** Zubin 211 **** '

      ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
      '*** Jay 3.2.0 [PicNm] (Added InCoCd)
      ' ******** Manali 3.03 - 18/07/08 - OdPoNo added
      ' **** Manali 3.6.0 - 12/11/09 - IdTrayNo added
      wCommStr = "Select InCoCd, InYy, InTc, InChr, InNo, max(InCmCd) as InCmCd, max(a.CmName) as InvCmName, " + _
               "max(InExpNo) as InExpNo, max(IdExpOdTc) as IdExpOdTc, max(IdExpOdYy) as IdExpOdYy, " + _
               "max(IdExpOdChr) as IdExpOdChr, max(IdExpOdNo) as IdExpOdNo, max(IdExpOdSr) as IdExpOdSr, " + _
               "max(OmCmCd) as OmCmCd, max(b.CmName) as OrdCmName, IdSr, max(IdTrayNo) as IdTrayNo, max(IdDmCd) as IdDmCd," + wOurDsgCd + "," + _
               "max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(IdSfx) as IdSfx, max(IdDmSz) as IdDmSz, max(OdDmCol) as OdDmCol," + _
               " max(OmPONo) as OmPONo, max(OdPoNo) as OdPoNo, IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= max(IdDmSz)), " + _
               " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= max(IdSfx) and CdOurDmSz= ''), " + _
               " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= max(IdDmSz)), " + _
               " IsNull((Select CdCustDmCd from CustDsg where CdCtg= 'C' and CdCd= max(OmCmCd) and CdOurDmCd= max(IdDmCd) and CdOurSfx= '' and CdOurDmSz= ''), '')))) as CdCustDmCd, " + _
               "max(IdKt) as IdKt, max(DmUom) as DmUom, max(IdGldLs) as IdGldLs, " + _
               "max(IdVaCtg) as IdVaCtg," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qDmKtDesc + "," + qVaCtgDesc + "," + _
               "(select RmCtg from RmMst join InvRm on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=IrPrtKey and ", "") + " RmCd=IrRmCd  and  " + _
               "IrMainMet='Y' Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " IrCocd = InCocd And IrTc = InTc  " + _
               "and IrYy = InYy and IrChr=InChr and IrNo=InNo and IrSr=IdSr) as OrRmCtg," + _
               "Max(" + wIdFOB + "/(case when IdQty> 0 then IdQty else 1 end)) as qUnitPrc, " + _
               "Max(" + wIdFOB + ") as IdVal, max(IdQty) as IdQty, "
      
    
  '***************Labour String***********************************
    
   Dim labstr, Labstr1 As String
     labstr = ""
     Labstr1 = ""
   If UCase(adc("UoLMCd1")) <> "" Then
      labstr = "sum(case when IlMcd='" + adc("UoLMcd1") + "' then isnull(" + wiLVal + ",0) else 0 end) as qLab1,"
   Else
      labstr = labstr + "0  as qLab1,"
   End If
   If UCase(adc("UoLMCd2")) <> "" Then
      labstr = labstr + "sum(case when IlMcd='" + adc("UoLMcd2") + "' then isnull(" + wiLVal + ",0) else 0 end) as qLab2,"
   Else
     labstr = labstr + "0  as qLab2,"
   End If
   If UCase(adc("UoLMCd3")) <> "" Then
       labstr = labstr + "sum(case when IlMcd='" + adc("UoLMcd3") + "' then isnull(" + wiLVal + ",0) else 0 end) as qLab3"
   Else
      labstr = labstr + "0  as qLab3"
   End If
      Labstr1 = "0  as qLab1," + "0  as qLab2," + "0  as qLab3"
          
    '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    wLabGrp = " group by InCoCd, InYy, InTc, InChr, InNo, IdSr"

    ' Zubin 211 (Patch 03 - Left Outer Join For InvLab and IsNull for wiLVal added)
    '****** Sachin 3.02 - Id Fields In Joins
    ' ***** Manali 3.5.0 - 11/11/08 - IrRmSz, qRmSzDesc, qIrLmeRt added Below [wLnSv]
    wLabCommStr = "'1LAB' as qTag,0 as qMainMet, 0 as IrSrNo, '' as RmCtg, '' as RmSCtg, " + _
                  "'' as IrRmCd, '' As RmQw,  0.0 as IrRmSz, '' as qRmSzDesc, " + _
                  " 0 as IrRmQty, 0 as IrRmIWt, 0 as IrRmIRt, 0 as qIrLmeRt, " + _
                  "'' as IrSetSCd, 0 as IrSetIRt, 0 as IrSetIVal, 0 as IrRmIVal," + _
                  " sum(IsNull(" + wiLVal + ", '')) as qTotLabVal , " + labstr + _
                  "  From InvHd " + _
                  " join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " And IdPrtKey=InPrtKey ", "") + _
                  " Left Outer join InvLab on IlIdIdNo=IdIdNo and iLCoCd= IdCoCd and iLTc= IdTc and iLYy= IdYy and " + _
                  " iLChr= IdChr and iLNo= IdNo and iLSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " and IlPrtKey=IdPrtKey ", "") + _
                  " join DsgMst on DmIdNo=IdDmIdNo and DmTcTyp='DM' and DmCd=IdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " And DmPrtKey=IdPrtKey ", "") + _
                  " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc " + _
                  " and OmYy= IdExpOdYy and OmChr= IdExpOdChr and OmNo = IdExpOdNo  " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=IdPrtKey ", "") + _
                  " join OrdDsg on OdCoCd=IdCoCd  and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
                  " and OdChr=IdExpOdChr and OdNo= IdExpOdNo and OdSr=IdExpOdSr  " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
                  " join CustMst a on a.CmCtg='C' and a.CmCd=InCmCd  " + _
                  " join CustMst b on b.CmCtg='C' and b.CmCd=OmCmCd  " + wCnd + wLabGrp
                  '(index= InvDsg) '(index= InvLab)

          
     
   '**********************************************************************
   
   
      ' urmi 2.04-3 RmRt when RmGrp <> Sr, and at RmCd and Lower level
      '****** Sachin 2.13.0 [18-08-06] [Q/W Module] <> 'W' changed to ='W' below
    wRtStr = " case when ((select RmQw from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=max(IrPrtKey) and ", "") + " RmCd= IrRmCd) = 'Q' and " + _
                          "sum(IrRmQty)<>0) then " + _
                          "sum(" + wIrRmVal + ") / sum(IrRmQty) " + _
                   "when ((select RmQw from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=max(IrPrtKey) and ", "") + " RmCd= IrRmCd) = 'W' and " + _
                          "sum(" + wIrRmWt + ")<>0) then " + _
                          "sum(" + wIrRmVal + ") / sum(" + wIrRmWt + ") " + _
                   "else 0 end "
 
  
  ' ****** Manali 3.5.0 - 11/11/08 - RmSzDesc [wLnSv]
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(Max(RisRmCtg), '') and PSCd= max(RmSCtg))= 'Y' then IsNull(Max(RisRmSzDesc), '') Else '' End "
  ' ****** Manali 3.5.0 - 11/11/08 - RmSzDesc

  '*************************Rm String  ***************************
    '****** Sachin 2.13.0 [18-08-06] - RmQw added below
    ' ***** Manali 3.5.0 - 11/11/08 - IrRmSz, qRmSzDesc, qIrLmeRt added Below
    Select Case UCase(adc("UoRmGrp"))
    Case Is = UCase("N")
      wRmStr = " IrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg, " + wOurRmCd + ",  max(IrRmSz) as IrRmSz, " + wRmSzDesc + " as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, max(" + wIrRmRt + ") as IrRmIRt, " + _
               "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as qIrLmeRt, " + _
               "max(IrSetSCd) as IrSetSCd, max(" + wIrSetRt + ") as IrSetIRt, sum(" + wIrSetVal + ") as IrSetIVal, " + _
               "sum(" + wIrRmVal + ") as IrRmIVal"
      wRmGrp = " Group by InPrtKey, InCoCd, InYy, InTc, InChr, InNo, IdSr, IrSrNo " '*** Jay 3.2.0 [PicNm] (Added InCoCd)

'      wRmStr = " IrSrNo, max(RmCtg), max(RmSCtg), max(IrRmCd), " + _
'               "sum(IrRmQty),sum(" + wIrRmWt + "), max(" + wIrRmRt + ")," + _
'               "max(" + wIrSetRt + "), sum(" + wIrSetVal + "), sum(" + wIrRmVal + ") "
'      wRmGrp = " group by InYy, InTc, InChr, InNo, IdSr, IrSrNo "
'
    Case Is = UCase("Sr")
      wRmStr = " IrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg," + wOurRmCd + ", max(IrRmSz) as IrRmSz, " + wRmSzDesc + " as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, max(" + wIrRmRt + ") as IrRmIRt, " + _
               "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as qIrLmeRt, " + _
               "max(IrSetSCd) as IrSetSCd, max(" + wIrSetRt + ") as IrSetIRt, sum(" + wIrSetVal + ") as IrSetIVal, " + _
               "sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, IrSrNo "     '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    Case Is = UCase("RmCdRt")
      wRmStr = " 0 as IrSrNo, RmCtg, RmSCtg, " + wOurRmCd + ", 0 as IrRmSz, '' as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, " + wIrRmRt + " as IrRmIRt, " + _
               "max(" + IIF(wb_InBaseCur, " (Case when InMulDiv= 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as qIrLmeRt, " + _
               "'' as IrSetSCd, 0 as IrSetIRt, sum(" + wIrSetVal + ") as IrSetIVal, sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, IrRmCd, " + wIrRmRt + " "   '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    Case Is = UCase("RmCd")
          wRmStr = " 0 as IrSrNo, RmCtg, RmSCtg, " + wOurRmCd + ", 0 as IrRmSz, '' as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, " + wRtStr + " as IrRmIRt, 0.00 as qIrLmeRt, '' as IrSetSCd, 0 as IrSetIRt, " + _
               "sum(" + wIrSetVal + ") as IrSetIVal, sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, IrRmCd "    '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    Case Is = UCase("RmSCtg")
      wRmStr = " 0 as IrSrNo, RmCtg, RmSCtg, '*' as IrRmCd, '*' As RmQw, 0 as IrRmSz, '' as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, 0 as IrRmIRt, 0.00 as qIrLmeRt, '' as IrSetSCd, " + _
               "0 as IrSetIRt, sum(" + wIrSetVal + ") as IrSetIVal, sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg "      '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    Case Is = UCase("RmCtg")
      wRmStr = " 0 as IrSrNo, RmCtg, '*' as RmSCtg, '*' as IrRmCd, '*' As RmQw, 0 as IrRmSz, '' as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, 0 as IrRmIRt, 0.00 as qIrLmeRt, '' as IrSetSCd, 0 as IrSetIRt, " + _
               "Sum (" + wIrSetVal + ") as IrSetIVal, Sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg "    '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    Case Is = UCase("RmRSzRt")
      wRmStr = " 0 as IrSrNo, max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg, max(IrRmCd) as IrRmCd, Max(RmQw) As RmQw, max(IrRmSz) as IrRmSz, " + wRmSzDesc + " as qRmSzDesc, " + _
               "sum(IrRmQty) as IrRmQty, sum(" + wIrRmWt + ") as IrRmIWt, max(" + wIrRmRt + ") as IrRmIRt, " + _
               "max (" + IIF(wb_InBaseCur, " (Case when InMulDiv = 'M' then " + wIrLmeRt + " /InCnvFct else " + wIrLmeRt + " *InCnvFct end)", wIrLmeRt) + ")  as qIrLmeRt, " + _
               "Max(IrSetSCd) as IrSetSCd, max(" + wIrSetRt + ") as IrSetIRt, " + _
               "sum (" + wIrSetVal + ") as IrSetIVal, Sum(" + wIrRmVal + ") as IrRmIVal "
      wRmGrp = " group by InPrtKey,InCoCd, InYy, InTc, InChr, InNo, IdSr, RmCtg, RmSCtg, RrRngSz "   '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    ' DispMsg "Option Available Only For Order Reports", etError: Exit Sub
    End Select
    
    '****** Sachin 3.02 - Id Fields in Joins
    ' ***** Manali 3.5.0 - 11/11/08 - Join with RmIdSz added Below [wLnSv]
    wRmCommStr = "'2RM' as qTag, (case when max(IrMainMet)= 'Y' or max(IrRmCd)= " + _
             "IsNull((Select IrRmCd From InvRm where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey=InPrtKey and ", "") + " IrCoCd= InCoCd and IrTc= InTc and IrYy= InYy and " + _
             "IrChr= InChr and IrNo= InNo and IrSr= IdSr and IrMainMet= 'Y'), '') then 1 " + _
             "else 2 end) as qMainMet, " + wRmStr + ",0 as qTotlabVal ," + Labstr1 + _
             " From InvHd " + _
             " join InvDsg on IdInIdNo=InIdNo and IdCoCd= InCoCd and IdTc= InTc and IdYy= InYy and IdChr= InChr and IdNo= InNo " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=InPrtKey ", "") + _
             " join InvRm on IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy " + _
             " and IrChr= IdChr and IrNo= IdNo and IrSr= IdSr " + IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey=IdPrtKey ", "") + _
             " join RmMst on RmCd= IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=IrPrtKey ", "") + _
             " join DsgMst on DmIdNo=IdDmIdNo And DmTcTyp='DM' and DmCd=IdDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=IdPrtKey ", "") + _
             " join OrdMst on OmCoCd= IdCoCd and OmTc= IdExpOdTc and  " + _
             " OmYy= IdExpOdYy and OmChr= IdExpOdChr and OmNo = IdExpOdNo  " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=IdPrtKey ", "") + _
             " join OrdDsg on OdCoCd=IdCoCd  and OdTc= IdExpOdTc and OdYy= IdExpOdYy " + _
             " and OdChr=IdExpOdChr and OdNo= IdExpOdNo and OdSr=IdExpOdSr  " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
             " join CustMst a on a.Cmctg='C' and a.CmCd=InCmCd  " + _
             " join CustMst b on b.Cmctg='C' and b.CmCd=OmCmCd " + _
             " join (Select OrPrtKey, OrOdIdNo, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd, max(OrLmeSal) as OrLmeSal from OrdRm " + _
             " Group By OrPrtKey, OrOdIdNo, OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd) OrdRm " + _
             " on OrCoCd=OdCoCd  and OrTc=  OdTc and OrYy=  OdYy  and OrChr= OdChr and OrNo= OdNo and OrSr=  OdSr   And OrPrtKey=OdPrtKey and OrRmCd=IrRmCd" + _
             " Left Outer join RmRt on RrTCTyp= 'RSZ' and RrCmCtg = 'C' and RrCmCd = 'ZSELF' and RrCmCurCd= '' and " + _
             " RrCtg = RmCtg and RrSCtg = RmSCtg and RrSTWGrp = '' and RrCd = '' and " + _
             " RrDmCtg = '' and RrLsCtg = '' and RrLabMCd = '' and RrFrLn <= IrRmSz and RrToLn >= IrRmSz " + IIF(gs_Partition = ctCurrPrtn, " and RrPrtKey=IrPrtKey ", "") + _
             " Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=IrRmSz " + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey='" + ctCurrPrtn + "' ", "") + wRmGrp
             '(index= InvDsg) '(index= InvRm)
   ' Debug.Print wCommStr + wRmCommStr
    wSqlStrg = wCommStr + wRmCommStr + " Union All " + _
               wCommStr + wLabCommStr + " order by InCoCd, InYy, InTc, InChr, InNo, IdSr, qTag,qMainMet,IrSrNo "  '*** Jay 3.2.0 [PicNm] (Added InCoCd)
    
    adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
        
    If adc("UoYN") = "N" Then
      Rep.PICSEC.Suppress = True
    'Else
    ' Rep.PicSec.Suppress = False
    End If
   Call DispCoNm
   ' **** Zubin 211 **** '
'''''        GRP_REP.SetFormula Rep, "wGrp3", "{@wInvNo}+ToText({rdo.IdSr}, '###0') " '{rdo.IdDmCd}+{rdo.IdDmCd}
        GRP_REP.SetFormula Rep, "wGrp3", wDetSrt + "+{@wInvNo}+ToText({rdo.IdSr}, '###0') " '{rdo.IdDmCd}+{rdo.IdDmCd}
  ' **** Zubin 211 **** '
        Rep.TxtLab1.SetText UCase(adc("UoLMCd1"))
        Rep.TxtLab2.SetText UCase(adc("UoLMCd2"))
        Rep.TxtLab3.SetText UCase(adc("UoLMCd3"))
        'Rep.PHDET.Suppress = False
    '*****************Suppress Rm  Details********************
    If adc("UoRmGrp") = "N" Then
      Rep.RmHead.Suppress = True
      Rep.RmDet.Suppress = True
      Rep.G3F.Suppress = True
      Rep.DcTotSec.Suppress = True
    End If
'    '****************Suppress Dia/Col Totals******************
    If adc("UoYN5") = "N" Then
       Rep.DcTotSec.Suppress = True
       Rep.hDiaCt.Suppress = True
    End If
    '*****************Suppress Po No & Customer Dsg Code*****************
    If adc("UoYN6") = "N" Then
        Rep.PoDesc.Suppress = True
    End If

  
'*******************Total Weights******************
    If adc("UoYN1") = "N" Then
     GRP_REP.SetFormula Rep, "wIrRmQty", "{rdo.IrRmQty}"
     GRP_REP.SetFormula Rep, "wIrRmVal", "{rdo.IrRmIVal}"
     GRP_REP.SetFormula Rep, "wRmWt", "{rdo.IrRmIWt}"
     GRP_REP.SetFormula Rep, "wSetVal", "{rdo.IrSetIVal}"
     GRP_REP.SetFormula Rep, "wLab1", "{rdo.qLab1}"
     GRP_REP.SetFormula Rep, "wLab2", "{rdo.qLab2}"
     GRP_REP.SetFormula Rep, "wLab3", "{rdo.qLab3}"
     GRP_REP.SetFormula Rep, "wTotLab", "{rdo.qTotLabVal}"
     GRP_REP.SetFormula Rep, "wDiaCt", " if Sum({@wDiaCol_Wt},{@wGrp3})<> 0 then " + _
     "{rdo.IdVal}/Sum({@wDiaCol_Wt},{@wGrp3}) else 0"
     Rep.Txtweights.SetText "(Total Weights)"
  ElseIf adc("UoYN1") = "Y" Then
  '************************Average  Weights******************
    GRP_REP.SetFormula Rep, "wIrRmQty", "{rdo.IrRmQty}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wIrRmVal", "{rdo.IrRmIVal}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wRmWt", "{rdo.IrRmIWt}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wSetVal", "{rdo.IrSetIVal}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wLab1", "{rdo.qLab1}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wLab2", "{rdo.qLab2}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wLab3", "{rdo.qLab3}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wTotLab", "{rdo.qTotLabVal}/{rdo.IdQty}"
    GRP_REP.SetFormula Rep, "wDiaCt", " if Sum({@wDiaCol_Wt},{@wGrp3})<> 0 then " + _
     "{rdo.IdVal}/({rdo.IdQty}*Sum({@wDiaCol_Wt},{@wGrp3})) else 0"
    Rep.Txtweights.SetText "(Average Weights)"
  End If
  
    If adc("UoYN5") = "N" Then
    GRP_REP.SetFormula Rep, "wDiaCt", " 0"
    End If
  
  '************Setting Group1 & Group2 Formula***********************
  If wGrpFormula1 <> "" Then ' Trim(adc("UoInGrp")) <> "" Then
    GRP_REP.SetFormula Rep, "wGrp1", wGrpFormula1
    GRP_REP.SetFormula Rep, "wGrp1Desc", wGrpDescFormula1
  End If
  If Trim(adc("UoInGrp1")) <> "" Then
    GRP_REP.SetFormula Rep, "wGrp2", wGrpFormula2
    GRP_REP.SetFormula Rep, "wGrp2Desc", wGrpDescFormula2
  End If
  If Trim(adc("UoInGrp")) = "" Then Rep.G1H.Suppress = True: Rep.G1F.Suppress = True
  If Trim(adc("UoInGrp1")) = "" Then Rep.G2H.Suppress = True: Rep.G2F.Suppress = True
     
     
  ' **** Manali 3.5.0 - 11/11/08 - wLnSv added in report
  GRP_REP.SetFormula Rep, "wLnSv", "If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                    "and Round({rdo.qIrLmeRt}, 3) <> 0.00  Then Trim (ToText ({rdo.qIrLmeRt}, 3)) " + _
                                    "else (If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} Else " + _
                                    "(If Round({rdo.IrRmSz}, 3)= 0.01 Then '+0 ' " + _
                                    "else if Round({rdo.IrRmSz}, 3)= 0.02 Then '+00 '  " + _
                                    "else if Round({rdo.IrRmSz}, 3)= 0.03 Then '+000 '  " + _
                                    "else if Round({rdo.IrRmSz}, 3)= 0 Then ' '  " + _
                                    "else Trim (ToText ({rdo.IrRmSz}, 3)) )) "
     
     '*************************Suppress Running Sr No or IdSr No********************
' **** Zubin 211 (UoYN3 and commented code below can be removed in Emr 2.12) **** '
'''''  If ADC("UoYN3") = "Y" Then
  ' **** Manali 3.6.0 - 12/11/09 - IdTrayNo [TRNO/SR] added
  If adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "DMSFX/SR" Or adc("UoInDet") = "TRNO/SR" Then
' **** Zubin 211 (UoYN3 and commented code above can be removed in Emr 2.12) **** '
   Rep.FldRunningSrNo.Suppress = False
   Rep.FldIdSr.Suppress = True
  Else
   Rep.FldRunningSrNo.Suppress = True
   Rep.FldIdSr.Suppress = False
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.IrRmCd}", 8)
  
  ' **** Manali 3.6.0 - 12/11/09 - IdTrayNo [TRNO/SR] added
  If adc("UoInDet") <> "TRNO/SR" Then Rep.TrayNo.Suppress = True
  
 CRV_REP.DisplayGroupTree = False
 
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******

End Sub

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPic"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "ShowAvgWt"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2     'Tag = "ShowVal"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 3     'Tag = "ShowRunSr"
    If .Value = Checked Then adc("UoYN3") = "Y" Else adc("UoYN3") = "N"
  Case Is = 4     'Tag = "ShowLabval"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  Case Is = 5     'Tag = "ShowDCValPerCts"
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  Case Is = 6     'Tag = "ShowCustPO"
    If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
  Case Is = 7     'Tag = "ShowCustDsg"
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  Case Is = 8     'Tag = "ShowCustRm"
    If .Value = Checked Then adc("UoYN8") = "Y" Else adc("UoYN8") = "N"
  Case Is = 9     'Tag = "InBaseCurr"
    If .Value = Checked Then
        adc("UoYN9") = "Y"
    Else
        If adc("UoInVal") = "S" Then
          .Value = Checked: adc("UoYN9") = "Y": Exit Sub
        Else
          adc("UoYN9") = "N"
        End If
    End If
  Case Is = 10     'Tag = "SortOrdCust"
    If .Value = Checked Then adc("UoYN10") = "Y" Else adc("UoYN10") = "N"
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
  Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
  Call InitProp(Me)
    If UCase(adc.MenuCd) = UCase("RepInvCustStmt") Then
    gs_RmGrp = "ORDDET"
    gs_LabGrp = "ORDDET"
    gs_PTyp = "LABMCD"
    gs_Tbl = "PARAM"
    Me.Caption = "Customer Statement"
  End If
  Call SetGroupSort

' **** Zubin 211 (Controls 'LblRunSr'/'UOYN3' and below code can be removed in Emr 2.12) **** '
  LblRunSr.Visible = False
  adc("UOYN3").Visible = False
  ChkBoxArr(3).Enabled = False: ChkBoxArr(3).Visible = False  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
' **** Zubin 211 (Controls 'LblRunSr'/'UOYN3' and above code can be removed in Emr 2.12) **** '

End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("UoYN2")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Show Value Option": Exit Sub


  Case Is = UCase("UoYN4")
      Cancel = pv_NewValue <> "" And (Not moCn.RecSeek("select PMCd from Param where " + _
                           " PTyp= 'YN' and PMCd = '" + pv_NewValue + "'"))
      ErrMsg = "Invalid Labour Value Option": Exit Sub
   End Select

  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
' Zubin 212 (Case for UoYN4 not required)
'  Case Is = UCase("UoYN4")
'    If adc("UoYN2") = "N" Then
'       Cancel = False
'    ElseIf adc("UoYN2") = "Y" Then
'      Cancel = True
'    End If
  ' Zubin 213
  Case Is = UCase("UoYN9")
    If UCase(adc("UoInVal")) = "S" Then Cancel = True: Exit Sub
    
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
 Select Case UCase(IdName)
' Zubin 212 (Recalc for UoYN4 not required)
'  Case Is = UCase("UoYN4")
'    If adc("UoYN2") = "Y" Then
'      adc("UoYN4") = ""
'    End If
  ' Zubin 213
  Case Is = UCase("UoYN9")
    If UCase(adc("UoInVal")) = "S" Then adc("UoYN9") = "Y": ChkBoxArr(9).Value = Checked: Exit Sub
  End Select
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UOUSRCD") = adc.UsrCd
    adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub DispCoNm()
'mRep.wCoCd.SetText gs_CoNm
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

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 adc("UoCoCdFr") = gs_CoCd
' **** Zubin 211 **** '
  If adc("UOINDET") = "" Then adc("UOINDET") = "INVSR"
' **** Zubin 211 **** '
  ' Zubin 212
  If adc("UOYN4") = "" Then adc("UOYN4") = "Y"
  ' Zubin 212
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN3") = "" Then adc("UoYN3") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  If adc("UoYN5") = "" Then adc("UoYN5") = "N"
  If adc("UoYN6") = "" Then adc("UoYN6") = "N"
  If adc("UoYN7") = "" Then adc("UoYN7") = "N"
  If adc("UoYN8") = "" Then adc("UoYN8") = "N"
  If adc("UoYN9") = "" Then adc("UoYN9") = "N"
  If adc("UoYN10") = "" Then adc("UoYN10") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN3") = "Y" Then ChkBoxArr(3).Value = Checked Else ChkBoxArr(3).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If adc("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If adc("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  If adc("UoYN10") = "Y" Then ChkBoxArr(10).Value = Checked Else ChkBoxArr(10).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
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
    Set mRep = Nothing
    Set Rep = Nothing
    '*** (09/08/05)
End Sub
