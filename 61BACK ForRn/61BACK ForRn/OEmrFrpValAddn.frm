VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrpValAddn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Value Addition"
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
      Left            =   11460
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   23
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
      Left            =   5565
      TabIndex        =   33
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   34
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
         TabIndex        =   35
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
      Left            =   6780
      TabIndex        =   32
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   29
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
      TabIndex        =   25
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
      TabIndex        =   28
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   27
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpValAddn.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpValAddn.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   31
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   26
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
            TabIndex        =   30
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Finding KT"
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
               TabIndex        =   21
               Tag             =   "ShowXBoe"
               ToolTipText     =   "Check To Show Finding KT in Details"
               Top             =   5100
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Finding Boe"
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
               Tag             =   "ShowXBoe"
               ToolTipText     =   "Check To Show Finding Boe in Report"
               Top             =   4815
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Profit %           "
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
               TabIndex        =   19
               Tag             =   "ShowProfit"
               ToolTipText     =   "Check To Show Profit (Show Report In DTA Format)"
               Top             =   4530
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Running Sr     "
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
               TabIndex        =   16
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Running Serial"
               Top             =   3960
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Hide Col 13-16          "
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
               TabIndex        =   14
               Tag             =   "HideCol13_16"
               ToolTipText     =   "Check To Hide Column 13 To 16"
               Top             =   3675
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Hide Col 12               "
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
               TabIndex        =   12
               Tag             =   "HideCol12"
               ToolTipText     =   "Check To Hide Column 12"
               Top             =   3390
               Width           =   2220
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Annexure Attached "
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
               TabIndex        =   10
               Tag             =   "Annex"
               ToolTipText     =   "Check To Show Annexure Attached"
               Top             =   3105
               Width           =   2220
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   20
               Left            =   2820
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   930
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IdYy ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   43
               Left            =   2040
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   930
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IdTc ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   45
               Left            =   3240
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   930
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IdChr ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   26
               Left            =   4020
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Number"
               Top             =   930
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IdNo ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   2040
               TabIndex        =   5
               ToolTipText     =   "Enter First Group"
               Top             =   1215
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP"
               IdName          =   "UOINGRP"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   2040
               TabIndex        =   6
               ToolTipText     =   "Enter Second Group"
               Top             =   1500
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINGRP1"
               IdName          =   "UOINGRP1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   11
               Left            =   2040
               TabIndex        =   8
               ToolTipText     =   "Specify Whether Finding Cost To Be Included Under Total Metal Cost Or Studding Cost"
               Top             =   2070
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   35
               Left            =   2040
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   645
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   2520
               TabIndex        =   15
               ToolTipText     =   "Show Running Serial Number or Invoice Serial  (Yes / No)"
               Top             =   3960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   2520
               TabIndex        =   9
               ToolTipText     =   "Annexure Attached (Y/N)"
               Top             =   3105
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   2040
               TabIndex        =   7
               ToolTipText     =   "Enter Detail Level To Be Shown In Report"
               Top             =   1785
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOINDET"
               IdName          =   "UOINDET"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   5
               Left            =   2520
               TabIndex        =   13
               ToolTipText     =   "Hide Columns 13 To 16"
               Top             =   3675
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   2520
               TabIndex        =   17
               ToolTipText     =   "Show Loss % (Yes / No)"
               Top             =   4245
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   2520
               TabIndex        =   11
               ToolTipText     =   "Hide Column 12"
               Top             =   3390
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               IdName          =   "UOYN4"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   2520
               TabIndex        =   18
               ToolTipText     =   "Display Report In Seepz (N) / DTA Format (Y)"
               Top             =   4530
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN5"
               IdName          =   "UOYN5"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   9
               Left            =   2520
               TabIndex        =   20
               ToolTipText     =   "Display Report In Seepz (N) / DTA Format (Y)"
               Top             =   4815
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN6"
               IdName          =   "UOYN6"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   2520
               TabIndex        =   51
               ToolTipText     =   "Display Finding Details in Report"
               Top             =   5100
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN7"
               IdName          =   "UOYN7"
            End
            Begin VB.Label LblProfitPer 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Profit % (Y/N)"
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
               TabIndex        =   50
               Top             =   4530
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label LblHideCol12 
               BackStyle       =   0  'Transparent
               Caption         =   "Hide Col 12"
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
               TabIndex        =   49
               Top             =   3390
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Loss %"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   48
               Top             =   4245
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.Label LblHideCols 
               BackStyle       =   0  'Transparent
               Caption         =   "Hide Cols 13-16"
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
               TabIndex        =   47
               Top             =   3675
               Visible         =   0   'False
               Width           =   1875
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
               TabIndex        =   46
               Top             =   1785
               Width           =   1845
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Annexure Attached"
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
               TabIndex        =   45
               Top             =   3105
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label LblRunnSr 
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
               TabIndex        =   44
               Top             =   3960
               Visible         =   0   'False
               Width           =   1875
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
               Index           =   1
               Left            =   90
               TabIndex        =   43
               Top             =   645
               Width           =   1335
            End
            Begin VB.Label LblFndCst 
               BackStyle       =   0  'Transparent
               Caption         =   "Finding Cost To Be Included With Metal Cost Or Studding Cost"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1035
               Left            =   90
               TabIndex        =   42
               Top             =   2070
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
               Left            =   90
               TabIndex        =   41
               ToolTipText     =   "Location"
               Top             =   1500
               Width           =   1875
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
               Left            =   90
               TabIndex        =   40
               ToolTipText     =   "Location"
               Top             =   1215
               Width           =   1875
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
               Left            =   2760
               TabIndex        =   39
               Top             =   930
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
               Left            =   3180
               TabIndex        =   38
               Top             =   930
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Tc/Yy/Chr/No"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   37
               Top             =   930
               Width           =   1815
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
               Index           =   0
               Left            =   3945
               TabIndex        =   36
               Top             =   930
               Width           =   105
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpValAddn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepValAddn
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

'*** Jen (2.12 Patch)
'Dim Rep1 As New EmrRepValAddnLocal   '*** (Bef 2.13)
Dim ws_FlagLocalExp As String
'Dim mRep As CRAXDRT.Report           '*** (Bef 2.13)
'*** Jen (2.12 Patch)
Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub

Private Sub adc_setreprecsource()
'*** (Jen 2.12)
On Error GoTo RepErr
'*** (Jen 2.12)

  '*** Report Sql ***
'****** (Jen 2.13) added fields qGldLsWt, qPlLsWt, qSlLsWt ******'
'*** Zubin 212 (Sales Ctg and qDmSalCtgDesc added)
'  Select 0 as qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, IrRmZWt as qIrWt,
'  IrRmZWt as qGrsWt, IrRmZWt as qNetWt, IrRmZWt as qGldWt, IrRmZWt as qPlWt, IrRmZWt as qSlWt, IrRmZWt as qPdWt,
'  IrRmZWt as qGldLsWt, IrRmZWt as qPlLsWt, IrRmZWt as qSlLsWt, IrRmZWt as qPdLsWt, IrRmZWt as qXWt, IrRmZWt as qDCSWt,
'  RmCtg, RmSCtg, IrRmZCd, RmDesc as qTyp, IrRmQty, IrRmZRt as qSalRt, space(22) as qIrZCdRt,
'  IrRmZVal as qDCSSalVal, IdZFob as qAvgSalPrc, IdZFob as qFOB, IdKt, IdVaCtg, DmCtg,
'  DmSalCtg, DmUom, OdDmCol, OrRmCtg, space(254) as qMainRmCtgDesc, space(254) as qDmCtgDesc,
'  space(254) as qDmSalCtgDesc, space(254) as qDmColDesc, space(254) as qVaCtgDesc, TBOENo,
'  TBOEDt, TzPurRt as qCstRt, IrRmZWt as qApportnWt
'  from InvDsg, OrdDsg, DsgMst, InvRm, OrdRm, RmMst, Txn, Txndz where 1= 2
  '*** Report Sql ***

  Dim wrepcnd As String, wCndA As String, wCndB As String, wStr1 As String, wStr2 As String
  Dim wStr3 As String, wStr4 As String, wStr5 As String, wSqlStrg1 As String, wSqlStrg2 As String
  Set Rep = Nothing
  
  Dim ws_CurCd As String  '*** (Jen 2.13)
  
  '*** (Bef 2.13)
  ''*** Jen (2.12 Patch)
  'Set Rep1 = Nothing
  'If ws_FlagLocalExp = "Y" Then
  '  Set adc.RepSource = Rep1
  '  Set mRep = Rep1
  'Else
  '  Set adc.RepSource = Rep
  '  Set mRep = Rep
  'End If
  ''*** Jen (2.12 Patch)
  '*** (Bef 2.13)
  
  Set adc.RepSource = Rep '*** (Bef 2.12 Patch)   '*** (Jen 2.13)
  
  Call DispCoNm
  ''Call GRP_REP.Gen3LRep
  ''grpflds = GRP_REP.GrpFldLst
  
  '*** For the Group Formula of the report
  Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
  Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
  Dim i As Integer, qDmCtgDesc As String, qMainRmCtgDesc As String, qDmColDesc As String
  Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String

' **** Zubin 211 **** '
  Dim wDetArr() As String, wDetSrt As String
' **** Zubin 211 **** '
  ' Zubin 212
  Dim qDmSalCtgDesc As String
  Dim ws_qIrWt As String    '*** (Jen 2.13)
  
  Dim ws_GrpBy As String
        
  qDmCtgDesc = "'' as qDmCtgDesc": qMainRmCtgDesc = "'' as qMainRmCtgDesc"
  qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc"
  qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
  
  ' Zubin 212
  qDmSalCtgDesc = "'' as qDmSalCtgDesc"
  wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp") + "' "), ",")
  For i = 0 To UBound(wGrpArr1)
    If UCase(wGrpArr1(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg"
      wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "OrRmCtg"
    ElseIf UCase(wGrpArr1(i)) = "KT" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.IdKt}"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdKt"
    ElseIf UCase(wGrpArr1(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "OdDmCol"
    ElseIf UCase(wGrpArr1(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdVaCtg"
    ElseIf UCase(wGrpArr1(i)) = "GLDLS" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim(ToText({rdo.IdGldLs}, '#0.00'))+ '%'"
      wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "}, '#0.00')"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdGldLs"
    ElseIf UCase(wGrpArr1(i)) = "PCS/PRS" Then
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr1(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "DmCtg"
    ' Zubin 212
    ElseIf UCase(wGrpArr1(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "DmSalCtg"
    ' Zubin 212
    End If
    'wGrpFormula1 = wGrpFormula1 + IIf(wGrpFormula1 = "", "", "+") + "Trim(UpperCase({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "}))"
    If wGrpArr1(i) <> "GLDLS" Then wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "}))"
  Next i
  
  wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
         "PMCd= '" + adc("UoInGrp1") + "' "), ",")
  For i = 0 To UBound(wGrpArr2)
    If UCase(wGrpArr2(i)) = "RMCTG" Then
      qMainRmCtgDesc = "(Select PDesc from Param where PTyp= 'RMCTG' and PMCd= max(OrRmCtg)) as qMainRmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qMainRmCtgDesc})"
      qMainMet = "max(OrRmCtg) as OrRmCtg"
      wOrJoin = " join OrdRm on " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OdPrtKey and ", "") + " OrOdIdNo=OdIdNo and OrCoCd= OdCoCd and OrTc= OdTc and OrYy= OdYy and OrChr= OdChr and OrNo= OdNo and OrSr= OdSr and OrMainMet= 'Y' "
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "OrRmCtg"
    ElseIf UCase(wGrpArr2(i)) = "KT" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.IdKt}"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdKt"
    ElseIf UCase(wGrpArr2(i)) = "DMCOL" Then
      qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= max(OdDmCol)) as qDmColDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "OdDmCol"
    ElseIf UCase(wGrpArr2(i)) = "VACTG" Then
      qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= max(IdVaCtg)) as qVaCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdVaCtg"
    ElseIf UCase(wGrpArr2(i)) = "GLDLS" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim(ToText({rdo.IdGldLs}, '#0.00'))+ '%'"
      wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "}, '#0.00')"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "IdGldLs"
    ElseIf UCase(wGrpArr2(i)) = "PCS/PRS" Then
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
    ElseIf UCase(wGrpArr2(i)) = "DMCTG" Then
      qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= max(DmCtg)) as qDmCtgDesc"
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "DmCtg"
    ' Zubin 212
    ElseIf UCase(wGrpArr2(i)) = "SALCTG" Then
      qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= max(DmCtg) and PSCd= max(DmSalCtg)) as qDmSalCtgDesc "
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      ws_GrpBy = ws_GrpBy + IIF(ws_GrpBy = "", "", ", ") + "DmSalCtg"
    ' Zubin 212
    End If
    If UCase(wGrpArr2(i)) <> "GLDLS" Then wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "}))"
  Next i
  '*** For the Group Formula of the report
  
' **** Zubin 211 **** '
  wDetArr = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INDET' and " + _
            "PMCd= '" + adc("UoInDet") + "' "), ",")
  For i = 0 To UBound(wDetArr)
    ' ***** Manali 3.8.0 - IdTrayNo in Sort
    wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", IIF(UCase(wDetArr(i)) = "IDTRAYNO", "ToText({rdo.IdTrayNo}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))"))
    ' bef 3.8.0
    'wDetSrt = wDetSrt + IIF(wDetSrt = "", "", "+") + IIF(UCase(wDetArr(i)) = "IDSR", "ToText({rdo.IdSr}, '###0')", "Trim(UpperCase({rdo." + wDetArr(i) + "}))")
  Next i
  
  If adc("UOINGRP") = "" Then GRP_REP.SetFormula Rep, "wGrp1", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  If adc("UOINGRP1") = "" Then GRP_REP.SetFormula Rep, "wGrp2", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  GRP_REP.SetFormula Rep, "wGrp3", wDetSrt + IIF(UCase(adc.MenuCd) = UCase("RepValAddnDta") And wDetSrt <> "", "+ ToText ({rdo.IdSr}, '###0')", "")
' **** Zubin 211 **** '

  ' ***** Manali 3.8.0 - Value Addn DTA Report
  If adc("UOINGRP") = "" Then SetSubFormula Rep.SubRep2, "wSubGrp1", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  If adc("UOINGRP1") = "" Then SetSubFormula Rep.SubRep2, "wSubGrp2", wDetSrt + IIF(wDetSrt <> "", "+", "") + " ToText ({rdo.IdSr}, '###0')"
  SetSubFormula Rep.SubRep2, "wSubGrp3", wDetSrt + IIF(UCase(adc.MenuCd) = UCase("RepValAddnDta") And wDetSrt <> "", "+ ToText ({rdo.IdSr}, '###0')", "")
  
  adc("UoCoCdFr").CmpStr = "M.TzCoCd = "
  adc("UoInTcFr").CmpStr = "M.TzTc= "
  adc("UoInYyFr").CmpStr = "M.TzYy= "
  adc("UoInChrFr").CmpStr = "M.TzChr= "
  adc("UoInNoFr").CmpStr = "M.TzNo= "
  wrepcnd = adc.RepCond
  wCndA = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  '*** Creating Temporary Table #TMPA ***
  '****** Sachin 3.02 - Id fields in Joins
  wStr1 = "Select IDENTITY(INT, 1, 1) as qTzSrNo, max(TBOENo) as TBOENo, " + _
          "max(TBOEDt) as TBOEDt, M.TzRmZCd as TzRmZCd, M.TzPurRt as qCstRt, " + _
          "M.TzRmZCd+ Space(12-Len(M.TzRmZCd))+ Str(IzRmZRt, 9, 2) as qTzZCdRt, " + _
          "IzRmZRt, sum(M.TzRmWt/(case when RmCtg in ('D', 'C') then 5 else 1 end)) as qTzWt " + _
          "into #TMPA From TxndZ M " + _
          "join TxndZ B on B.TzCoCd= M.TzCoCd and B.TzYy= M.TzRefYy and B.TzKey= M.TzRefKey and B.TzBM= 'B' " + IIF(gs_Partition = ctCurrPrtn, " and B.TzPrtKey=M.TzPrtKey ", "") + _
          "join Txn on TIdNo=B.TzTIdNo and TCoCd= B.TzCoCd and TTc= B.TzTc and TYy= B.TzYy " + _
          "and TChr= B.TzChr and TNo= B.TzNo " + IIF(gs_Partition = ctCurrPrtn, " and TPrtKey=B.TzPrtKey ", "") + _
          "join InvZRm on IzCoCd= M.TzCoCd and IzTc= M.TzTc and IzYy= M.TzYy and IzChr= M.TzChr and " + _
          "IzNo= M.TzNo and IzSr= M.TzSr " + IIF(gs_Partition = ctCurrPrtn, " and IzPrtKey=M.TzPrtKey ", "") + " join RmMst on RmCd= M.TzRmZCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=M.TzPrtKey ", "") + _
          wCndA + IIF(gs_Partition = ctCurrPrtn, " and M.TzPrtKey='" + ctCurrPrtn + "' ", "") + _
          " Group By TBOENo, M.TzRmZCd, IzRmZRt, M.TzPurRt "
  moCn.CreateTmpTable "#TMPA", wStr1

  adc("UoCoCdFr").CmpStr = "IdCoCd = "
  adc("UoInTcFr").CmpStr = "IdTc= "
  adc("UoInYyFr").CmpStr = "IdYy= "
  adc("UoInChrFr").CmpStr = "IdChr= "
  adc("UoInNoFr").CmpStr = "IdNo= "
  wrepcnd = adc.RepCond
  wCndB = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '*** (Bef 2.13)
  '"sum(IrRmZWt/(case when Rm.RmCtg in ('D', 'C') then 5 else 1 end)* " + _
   "(case when Rm.RmCtg in ('G', 'P', 'S') then Rm.RmPurityZ* (1 + IdGldLs/100) else 1 end)) as qIrWt "
  '*** (Bef 2.13)

  '*** (Jen 2.13) ***
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L' added
  If moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "'") = "Y" Then
    ws_qIrWt = " sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "    / (case when Rm.RmCtg in ('D', 'C') then 5 else 1 end) " + _
               "    * (case when Rm.RmCtg in ('G', 'P', 'S', 'L','M') then Rm.RmPurityZ" + _
               "            * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/ 100) " + _
               "            / (case when Z.RmPurityZ> 0 then Z.RmPurityZ else 1 end) else 1 end) " + _
               "    / (case when (Rm.RmCtg= 'G' and InHGldFiness> 0) then InHGldFiness " + _
               "            when (Rm.RmCtg= 'P' and InHPlFiness> 0) then InHPlFiness " + _
               "            when (Rm.RmCtg= 'S' and InHSlFiness> 0) then InHSlFiness " + _
               "            when (Rm.RmCtg= 'L' and InHPdFiness> 0) then InHPdFiness else 1 end) " + _
               "   ) as qIrWt "
  Else
    ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L' added
    ws_qIrWt = " sum((case when IrMainMet= 'Y' then IdCustmMtchWt else IrRmZWt end) " + _
               "    / (case when Rm.RmCtg in ('D', 'C') then 5 else 1 end) " + _
               "    * (case when Rm.RmCtg in ('G', 'P', 'S', 'L','M') then Rm.RmPurityZ" + _
               "            * (1 + (case when IrMainMet= 'Y' then 0 else IdGldLs end)/100) " + _
               "            / (case when Z.RmPurityZ> 0 then Z.RmPurityZ else 1 end) else 1 end) " + _
               "   ) as qIrWt "
  End If
  '*** (Jen 2.13) ***

  '*** (Bef 2.13)
  '"sum(IrRmQty*" + IIF(UCase(adc("UoPMCdFr")) = "STUDCST", "1", "(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)") + ") as IrRmQty, "
  '*** (Bef 2.13)

  '****** (Jen 2.13) added fields qGldLsWt, qPlLsWt, qSlLsWt. Also InvHd join added and changed IrRmQty ******'
  ' Zubin 212 (DmSalCtg added)
  '****** Sachin 3.02 - Id fields in Joins
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  
  Dim wMetWtStr As String, wLsWtStr As String
  
  
  If UCase(adc.MenuCd) = UCase("RepStudStmt") Then
    wMetWtStr = "IrRmZWt*(case when Rm.RmCtg = 'M' then 1 else (IsNull(x1.RmPurityZ,Rm.RmPurityZ)/(case when z.RmPurityZ> 0 then z.RmPurityZ else 1 end)) end)"
    wLsWtStr = "(IdCustmMtchWt- IrRmZWt)*Rm.RmPurityZ/(case when z.RmPurityZ> 0 then z.RmPurityZ else 1 end)"
    Rep.TxtRepTitle.SetText ("STATEMENT OF STUDDING PARTICULARS")
    
    Rep.FldXApportnWt.Suppress = True
    Rep.gXWt.Suppress = False
    
    Rep.Fldg2GWt.Left = Rep.Fldg2GWt.Left + 1500: Rep.Fldg2PWt.Left = Rep.Fldg2PWt.Left + 1500
    Rep.Fldg2SlWt.Left = Rep.Fldg2SlWt.Left + 1500: Rep.Fldg2PdWt.Left = Rep.Fldg2PdWt.Left + 1500
    Rep.g1NetPureWt.Left = Rep.g1NetPureWt.Left + 1500: Rep.gNetPureWt.Left = Rep.gNetPureWt.Left + 1500
    
    Rep.g2GLossWt1.Left = Rep.g2GLossWt1.Left + 630: Rep.g2PLossWt1.Left = Rep.g2PLossWt1.Left + 630
    Rep.g2SLossWt1.Left = Rep.g2SLossWt1.Left + 630: Rep.g2LLossWt1.Left = Rep.g2LLossWt1.Left + 630
    Rep.gLossWt.Left = Rep.gLossWt.Left + 630: Rep.g1LossWt.Left = Rep.g1LossWt.Left + 630
    Rep.g2GLossPrc.Suppress = True: Rep.g2PLossPrc.Suppress = True: Rep.g2SLossPrc.Suppress = True: Rep.g2LLossPrc.Suppress = True
    
  Else
    wMetWtStr = "IrRmZWt"
    wLsWtStr = "(IdCustmMtchWt- IrRmZWt)"
    Rep.TxtRepTitle.SetText ("VALUE ADDITION")
    
    Rep.FldXApportnWt.Suppress = False
    Rep.gXWt.Suppress = True
    
     Rep.g2GNetWt.Suppress = True: Rep.g2PlNetWt.Suppress = True: Rep.g2SlNetWt.Suppress = True
    Rep.g2PdNetWt.Suppress = True
    Rep.g1NetPureWt.Suppress = True: Rep.gNetPureWt.Suppress = True:
    
  End If
  
  wStr2 = "Select IDENTITY(INT, 1, 1) AS qIrSrNo, IdSr, max(IdDmCd) as IdDmCd, " + _
          "max(IdDmSz) as IdDmSz, max(IdQty) as IdQty, max(IdGldLs) as IdGldLs, " + ws_qIrWt + ", " + _
          "sum(IrRmZWt/(case when Rm.RmCtg in ('D', 'C') then 5 else 1 end)) as qGrsWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then 1 else 0 end)) as qNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdNetWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('M') then 1 else 0 end)) as qMNetWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('G', 'P', 'S', 'L', 'M') then 1 else 0 end)) as qNetPureWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('M') then 1 else 0 end)) as qMWt, "

  wStr2 = wStr2 + "sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('G') then 1 else 0 end)) as qGldLsWt, " + _
          "sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('P') then 1 else 0 end)) as qPlLsWt, " + _
          "sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('S') then 1 else 0 end)) as qSlLsWt, " + _
          "sum((case when IrMainMet= 'Y' then " + wLsWtStr + " else (" + wMetWtStr + "* IdGldLs/100) end)* (case when Rm.RmCtg in ('L') then 1 else 0 end)) as qPdLsWt, " + _
          "0 as qMLsWt, " + _
          "sum(" + wMetWtStr + "*(case when Rm.RmCtg in ('X') then 1 else 0 end)) as qXWt, " + _
          "sum(IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end))  as qDCSWt, " + _
          "sum(IrRmQty*(case when Rm.RmCtg in ('X') then 1 else 0 end)) as qXIrRmQty, " + _
          "max(Rm.RmCtg) as RmCtg, max(Rm.RmSCtg) as RmSCtg, IrRmZCd, max(Z.RmDesc) as qTyp, " + _
          "sum(IrRmQty*" + IIF(UCase(adc("UoPMCdFr")) = "STUDCST", "(case when Rm.RmCtg in ('D', 'C', 'X') then 1 else 0 end)", "(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)") + ") as IrRmQty, " + _
          "IrRmZCd+ Space(12-Len(IrRmZCd))+ Str(IrRmZPureRt, 9, 2) as qIrZCdRt, IrRmZPureRt as qSalRt, " + _
          "sum(IrRmZRt*IrRmZWt*(case when Rm.RmCtg in ('D', 'C') then 1 else 0 end)) as qDCSSalVal, " + _
          "max(case when IdQty> 0 then IdZFob/IdQty else 0 end) as qAvgSalPrc, " + _
          "max(IdZFob) as qFOB, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, " + _
          "max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, " + qMainMet + ", " + _
          "max(IdTrayNo) as IdTrayNo "
            
  '****** Sachin 3.02 - Id fields in Joins
  wStr2 = wStr2 + " into #TMPB " + _
          "From InvRm join RmMst Rm on " + IIF(gs_Partition = ctCurrPrtn, " Rm.RmPrtKey=IrPrtKey and ", "") + " Rm.RmCd= IrRmCd join RmMst Z on " + IIF(gs_Partition = ctCurrPrtn, " Z.RmPrtKey=IrPrtKey and ", "") + _
          " Z.RmCd= (Case when Rm.RmCtg='X' Then (Select RmZCd From RmMst x where x.RmCd=Rm.RmBaseCd) Else  IrRmZCd End) " + _
          "join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
          "join InvHd on InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IdPrtKey ", "") + _
          "join OrdDsg on OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy and " + _
          "OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=IdPrtKey ", "") + _
          "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + wOrJoin + _
          "Left Outer Join RmMst x1 On x1.RmCd=Rm.RmBaseCd " + _
          "Where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " Round(IrRmZWt, 3) > 0 " + wCndB + " Group By IdSr, IrRmZCd, IrRmZPureRt Order By IdSr, IrRmZCd, qSalRt "
  moCn.CreateTmpTable "#TMPB", wStr2
  
  '*** Creating Temporary Table #TMPC ***
  ' Zubin 212 (DmSalCtg added)
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - Pd fields added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  wStr3 = "Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt,  " + _
          "qGldWt, qPlWt, qSlWt, qPdWt, qMWt, qNetPureWt, qGldNetWt, qPlNetWt, qSlNetWt, qPdNetWt, qMNetWt, qGldLsWt, qPlLsWt, qSlLsWt, qPdLsWt, qMLsWt, qXWt, qDCSWt, qXIrRmQty, RmCtg, RmSCtg, IrRmZCd, " + _
          "qTyp, IrRmQty, qIrZCdRt, qSalRt, qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, " + _
          "DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, qTzSrNo, TBOENo, TBOEDt, qCstRt, qTzWt  " + _
          "into #TMPC " + _
          "From #TMPB join #TMPA on qTzZCdRt= qIrZCdRt order By qIrSrNo, qTzWt, qTzSrNo "
  moCn.CreateTmpTable "#TMPC", wStr3

  Call ApportnTbl
  
  '*** Inserting Recs in Temporary Table #TMPD which is the main Apportioning Sql ******
  ' Zubin 212 (DmSalCtg added)
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - Pd Fields added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  wStr4 = "Insert into #TMPD Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, " + _
          "qGldWt, qPlWt, qSlWt, qPdWt, qMWt, qNetPureWt, qGldNetWt, qPlNetWt, qSlNetWt, qPdNetWt, qMNetWt, qGldLsWt, qPlLsWt, qSlLsWt, qPdLsWt, qMLsWt, qXWt, qDCSWt, qXIrRmQty, RmCtg, RmSCtg, IrRmZCd, qTyp, " + _
          "IrRmQty, qSalRt, qIrZCdRt, qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, " + _
          "OdDmCol, OrRmCtg, IdTrayNo, TBOENo, TBOEDt, qCstRt, " + _
          "qIrWt- (case when IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) > 0 " + _
          "             then IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) " + _
          "             else 0 end) as qFinalReq, "
  wStr5 = "qTzWt- (case when IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) > 0 " + _
          "             then IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
          "                          and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
          "                  IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
          "                          and a.qTzZCdRt= c.qIrZCdRt), 0) " + _
          "             else 0 end) as qFinalAvl from #TMPC c "
  
  moCn.Execute (wStr4 + wStr5)
  ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - Pd fields added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  wSqlStrg1 = "Select 0 As qIrSrNo, IdSr, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, " + _
              "max(IdQty) as IdQty, max(IdGldLs) as IdGldLs, 0 as qIrWt, " + _
              "sum(qGrsWt) as qGrsWt, sum(qNetWt) as qNetWt, sum(qGldWt) as qGldWt, " + _
              "sum(qPlWt) as qPlWt, sum(qSlWt) as qSlWt, sum(qPdWt) as qPdWt, sum(qMWt) as qMWt, sum(qNetPureWt) as qNetPureWt, " + _
              "sum(qGldNetWt) as qGldNetWt, sum(qPlNetWt) as qPlNetWt, sum(qSlNetWt) as qSlNetWt, sum(qPdNetWt) as qPdNetWt, sum(qMNetWt) as qMNetWt, sum(qGldLsWt) as qGldLsWt, " + _
              "sum(qPlLsWt) as qPlLsWt, sum(qSlLsWt) as qSlLsWt, sum(qPdLsWt) as qPdLsWt,sum(qMLsWt) as qMLsWt, sum(qXWt) as qXWt, sum(qDCSWt) as qDCSWt, Sum(qXIrRmQty) as qXIrRmQty, " + _
              "'' as RmCtg, '' as RmSCtg, '' as IrRmZCd, '' as qTyp, " + _
              "sum(IrRmQty) as IrRmQty, 0 as qSalRt, '' as qIrZCdRt, " + _
              "sum(qDCSSalVal) as qDCSSalVal, max(qAvgSalPrc) as qAvgSalPrc, " + _
              "max(qFOB) as qFOB, " + _
              "max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, " + _
              "max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, max(OrRmCtg) as OrRmCtg, max(IdTrayNo) as IdTrayNo, " + _
              qMainRmCtgDesc + ", " + qDmCtgDesc + ", " + qDmSalCtgDesc + ", " + qDmColDesc + ", " + qVaCtgDesc + ", " + _
              "'' as TBOENo, '' as TBOEDt, 0 as qCstRt, 0 as qApportnWt " + _
              "From #TMPB Group By IdSr "

  ' Zubin 212 (DmSalCtg and qDmSalCtgDesc added)
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - Pd fields added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  wSqlStrg2 = "Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, " + _
              "qGldWt, qPlWt, qSlWt, qPdWt, qMWt, qNetPureWt, qGldNetWt, qPlNetWt, qSlNetWt, qPdNetWt, qMNetWt, qGldLsWt, qPlLsWt, qSlLsWt, qPdLsWt,qMLsWt, qXWt, qDCSWt, 0 as qXIrRmQty, RmCtg, RmSCtg, " + _
              "IrRmZCd, qTyp, IrRmQty, qSalRt, qIrZCdRt, qDCSSalVal, qAvgSalPrc, qFOB, IdKt, " + _
              "IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, '' as qMainRmCtgDesc, " + _
              "'' as qDmCtgDesc, '' as qDmSalCtgDesc, '' as qDmColDesc, '' as qVaCtgDesc, " + _
              "TBOENo, TBOEDt, qCstRt, " + _
              "(case when qFinalReq>0 and qFinalAvl> 0 " + _
              "      then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) " + _
              "      else 0 end) as qApportnWt from #TMPD "
              
  ' ***** Manali 3.9.0 - Invoice Rm-Fgd Mismatch
  Dim ws_IdIfMisMatch As String
  ws_IdIfMisMatch = "Select IdSr From InvDsg " + _
                "left outer join " + _
                "(Select IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr, " + _
                "IsNull(sum(IrRmAWt/(case when IrRmCtg in ('D','C') then 5 else 1 End)), 0) as qIrGrsWt from InvRm " + _
                "Group By IrPrtKey, IrIdIdNo, IrCoCd, IrTc, IrYy, IrChr, IrNo, IrSr) Ir " + _
                "on IdCoCd=IrCoCd and IdTc=IrTc and IdYy=IrYy and IdChr=IrChr and IdNo=IrNo and IdSr=IrSr and IdPrtKey=IrPrtKey and IdIdNo=IrIdIdNo" + _
                " left outer join " + _
                "(Select IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr, IsNull(sum(IfGrsWt), 0) as qIfGrsWt, IsNull(sum(IfQty), 0) as qIfQty from InvFgd " + _
                "Group By IfPrtKey, IfCoCd, IfTc, IfYy, IfChr, IfNo, IfSr) Inf " + _
                "on IrCoCd=IfCoCd and IrTc=IfTc and IrYy=IfYy and IrChr=IfChr and IrNo=IfNo and IrSr=IfSr and IrPrtKey=IfPrtKey " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " IdPrtKey='" + ctCurrPrtn + "' and ", "") + _
                " IdCoCd= '" + adc("UoCoCdFr") + "' and " + _
                "IdTc= '" + adc("UoInTcFr") + "' and IdYy= '" + adc("UoInYyFr") + "' and IdChr= '" + adc("UoInChrFr") + "' and " + _
                "IdNo= " + CStr(adc("UoInNoFr")) + " " + _
                "and (abs(cast(cast(IsNull(qIrGrsWt, 0) as decimal(14,3)) as decimal(14,2)) - cast(cast(IsNull(qIfGrsWt, 0) as decimal (14,3)) as decimal(14,2))) > 0.9 " + _
                " or cast(IsNull(qIfQty ,0) as decimal(14,0)) <> cast(IsNull(IdQty ,0) as decimal (14,0))) "
  If moCn.RecSeek(ws_IdIfMisMatch) Then
    adc.RepRecSource = "Select * From #TmpD Where 1=2" + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
    DispMsg "Invoice Contains Discrepancy, Cannot Display Report.", etError
    Exit Sub
  End If
  
  '****** Sachin 2.14.0 Next - [20-02-2007] ******
  '****** Sachin 3.02 - Id fields in Joins
  ' ******** Manali 19/09/08 - [IrRmZCd<>'' condition added in 2nd query]
  If moCn.RecSeek("Select 'x' From InvZRm " + _
       " Left Outer Join Txndz On TzCoCd=IzCoCd And TzTc=IzTc And TzYy=IzYy And TzChr=IzChr And TzNo=IzNo And TzSr=IzSr " + _
       IIF(gs_Partition = ctCurrPrtn, " and TzPrtKey= IzPrtKey ", "") + _
       " Where " + IIF(gs_Partition = ctCurrPrtn, " IzPrtKey='" + ctCurrPrtn + "' and ", "") + _
       " IzCoCd= '" + adc("UoCoCdFr") + "' And " + _
       " IzTc= '" + adc("UoInTcFr") + "' and IzYy= '" + adc("UoInYyFr") + "' and " + _
       " IzChr= '" + adc("UoInChrFr") + "' and IzNo= " + CStr(adc("UoInNoFr")) + _
       " Group By IzCoCd, IzTc, IzYy, IzChr, IzNo, IzSr " + _
       " Having Sum(convert(decimal (16, 3), IsNull(TzRmWt, 0)))- Max(convert(decimal (16, 3), IsNull(IzRmZWt, 0))) <> 0") _
     Or _
     moCn.RecSeek("Select 'x' from InvRm " + _
                "where IrCoCd= '" + adc("UoCoCdFr") + "' and IrTc= '" + adc("UoInTcFr") + "' " + _
                "and IrYy= '" + adc("UoInYyFr") + "' and IrChr= '" + adc("UoInChrFr") + "' " + _
                "and IrNo= " + CStr(adc("UoInNoFr")) + _
                IIF(gs_Partition = ctCurrPrtn, " and IrPrtKey= '" + ctCurrPrtn + "' ", "") + _
                "and convert(decimal (16, 3), IrRmZWt)<>0 and IrRmZCd<>'' and not exists(Select 'x' from InvZRm where IzCoCd= IrCoCd and IzTc= IrTc " + _
                "and IzYy= IrYy and IzChr= IrChr and IzNo= IrNo " + _
                IIF(gs_Partition = ctCurrPrtn, " and IzPrtKey= IrPrtKey ", " ") + _
                "and IzRmZCd= IrRmZCd and convert(decimal(14, 3), IzRmZRt)= convert(decimal(14, 3), IrRmZPureRt)) ") Then
                 
    adc.RepRecSource = "Select * From #TmpD Where 1=2" + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
    DispMsg "Incomplete Matching Done, Cannot Display Report.", etError
    Exit Sub
  Else
    adc.RepRecSource = wSqlStrg1 + " Union All " + wSqlStrg2 + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  End If
  '****** Sachin 2.14.0 Next - [20-02-2007] ******
  
  '*** (Jen 2.12)
  Dim ws_SubRep As String
  ws_SubRep = "Select IsNull(PSCd, 'ZClrStn') as qStnCls, " + _
              "IsNull(max(PDesc), 'Color Stones') as qStnClsDesc, sum(qTzWt* 5) as qWt, " + _
              "sum(qTzWt* 5 * qCstRt) as qCst " + _
              "From #TMPA join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= TzRmZCd " + _
              "left outer join Param on PTyp= 'STNCLS' and PmCd= RmCtg " + _
              "   and ',' + PDesc225 + ',' like '%,'+ RmSCtg +',%'" + _
              "where RmCtg= 'C' and RmZ= 'Y' " + _
              "Group By PSCd Order By qStnCls "
  Call moCn.RepRes(Rep.Subreport1.OpenSubreport, ws_SubRep + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  '*** (Jen 2.12)
              
  '3.11.0
  Dim ws_XRmKt As String
  ws_XRmKt = "Select Distinct IdSr, R.RmKt From #TMPD Join RmMst R On R.RmZ='Y' And R.RmCd=IrRmZCd Where R.RmCtg='X'"
  Call moCn.RepRes(Rep.SubRepXRmKt.OpenSubreport, ws_XRmKt + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
             
              
''  '*** Apportioning Sql ***
''  wSqlStrg2 = "Select qIrSrNo, IdSr, IdDmCd, IdDmSz, IdQty, IdGldLs, qIrWt, qGrsWt, qNetWt, " + _
''              "qGldWt, qPlWt, qXWt, qDCSWt, RmCtg, RmSCtg, IrRmZCd, qTyp, IrRmQty, qSalRt, qIrZCdRt, " + _
''              "qDCSSalVal, qAvgSalPrc, qFOB, IdKt, IdVaCtg, DmCtg, DmUom, OdDmCol, OrRmCtg, " + _
''              "'' as qMainRmCtgDesc, '' as qDmCtgDesc, '' as qDmColDesc, '' as qVaCtgDesc, " + _
''              "TBOENo, TBOEDt, qCstRt, " + _
''              "(case when (Select sum(qIrWt) from #TMPB b where b.qIrSrNo<= c.qIrSrNo " + _
''              "            and b.qIrZCdRt= c.qIrZCdRt) - " + _
''              "           IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
''              "                   and a.qTzZCdRt= c.qIrZCdRt), 0)> 0 " + _
''              "           and IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
''              "                       and b.qIrZCdRt= c.qIrZCdRt), 0) - " + _
''              "           IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo<= c.qTzSrNo " + _
''              "                   and a.qTzZCdRt= c.qIrZCdRt), 0)<= 0 "
''  wSqlStrg3 = "     then (case when IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo<= c.qTzSrNo " + _
''              "                              and a.qTzZCdRt= c.qIrZCdRt), 0)- " + _
''              "                      IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
''              "                              and b.qIrZCdRt= c.qIrZCdRt), 0) > " + _
''              "                     (Select sum(qIrWt) from #TMPB b where b.qIrSrNo<= c.qIrSrNo " + _
''              "                              and b.qIrZCdRt= c.qIrZCdRt) - " + _
''              "                     IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
''              "                             and a.qTzZCdRt= c.qIrZCdRt), 0) " + _
''              "                 then (Select sum(qIrWt) from #TMPB b where b.qIrSrNo<= c.qIrSrNo " + _
''              "                       and b.qIrZCdRt= c.qIrZCdRt) - " + _
''              "                      IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo< c.qTzSrNo " + _
''              "                              and a.qTzZCdRt= c.qIrZCdRt), 0) " + _
''              "                 else IsNull((Select sum(qTzWt) from #TMPA a where a.qTzSrNo<= c.qTzSrNo " + _
''              "                              and a.qTzZCdRt= c.qIrZCdRt), 0)- " + _
''              "                      IsNull((Select sum(qIrWt) from #TMPB b where b.qIrSrNo< c.qIrSrNo " + _
''              "                              and b.qIrZCdRt= c.qIrZCdRt), 0) end) " + _
''              "      else 0 end) as qApportnWt from #TMPC c "
''
''  ADC.RepRecSource = wSqlStrg1 + " Union All " + wSqlStrg2 + wSqlStrg3
  
  Dim wRsInHd As MDORowSet
  Dim w_InHValAddMtlColYN As String
  Dim w_InComPer As Double
  
  '*** (Bef 2.13)
  'Set wRsInHd = moCn.OpenRes("Select InTc, InYy, InChr, InNo, InPrnDt, InExpNo, InSelfName, " + _
  '              "InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InVaDC, InVaX, " + _
  '              "InCnvRt, InGrNo, InGrDt, InHValAddnFt, InhAuthSgn, InHValAddMtlColYN, InComPer " + _
  '              "from InvHd where InCoCd ='" + adc("UoCoCdFr") + "' and InTc= '" + adc("UoInTcFr") + "' and " + _
  '              "InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "' and " + _
  '              "InNo= " + CStr(adc("UoInNoFr")) + " ")
  '*** (Bef 2.13)
  
  '*** (Jen 2.13)
  ' ***** Manali 3.6.0 - 17/11/09 - InCircularNo, InFob added
  Set wRsInHd = moCn.OpenRes("Select InTc, InYy, InChr, InNo, InPrnDt, InExpNo, InSelfName, " + _
                "InSelfAdd1, InSelfAdd2, InSelfAdd3, InSelfTel, InSelfFax, InVaDC, InVaX, " + _
                "(case when InRsMulDiv= 'D' and convert (decimal (16, 4), InCnvRt)> 0 then 1/ InCnvRt else InCnvRt end) as InCnvRt, " + _
                "InGrNo, InGrDt, InHValAddnFt, InhAuthSgn, InHValAddMtlColYN, InComPer, CmCurCd, InCircularNo, InFob " + _
                "from InvHd join Param on PTyp= 'TC' and PMCd= InTc and PSCd= '' " + _
                "join CustMst on CmCtg= (case PValue when 'IN' then 'C' when 'CTB' then 'T' else '' end) and CmCd= InCmCd " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd ='" + adc("UoCoCdFr") + "' and InTc= '" + adc("UoInTcFr") + "' and " + _
                "InYy= '" + adc("UoInYyFr") + "' and InChr= '" + adc("UoInChrFr") + "' and " + _
                "InNo= " + CStr(adc("UoInNoFr")) + " ")
  '*** (Jen 2.13)
  
  Dim wrs_CompanyRecord As MDORowSet
  Dim ws As String
  ws = "Select CmBuyAdd1, CmBuyAdd2, CmBuyAdd3 from CustMst where CmCtg='Z' and CmCd='" + adc("UoCoCdFr") + "' "
  Set wrs_CompanyRecord = moCn.OpenRes(ws)
  If wrs_CompanyRecord.RecCount > 0 Then
    SetRepText Rep, "TxtRegOff", Trim(wrs_CompanyRecord!CmBuyAdd1) + " " + Trim(wrs_CompanyRecord!CmBuyAdd2) + _
                                 " " + Trim(wrs_CompanyRecord!CmBuyAdd3)
  End If
 


  If wRsInHd.RecCount > 0 Then
'    SetRepText Rep, "TxtRegOff", Trim(wRsInHd!InSelfAdd1) + " " + Trim(wRsInHd!InSelfAdd2) + _
                                 " " + Trim(wRsInHd!InSelfAdd3)
    SetRepText Rep, "TxtFact", Trim(wRsInHd!InSelfAdd1) + " " + Trim(wRsInHd!InSelfAdd2) + _
                                 " " + Trim(wRsInHd!InSelfAdd3)
    SetRepText Rep, "TxtRegOffTel", wRsInHd!InSelfTel
    SetRepText Rep, "TxtRegOffFax", wRsInHd!InSelfFax
    SetRepText Rep, "TxtFactTel", wRsInHd!InSelfTel
    SetRepText Rep, "TxtFactFax", wRsInHd!InSelfFax
    SetRepText Rep, "TxtInvVchNo", Trim(wRsInHd!InTc) + "/" + Trim(wRsInHd!InYy) + "/" + Trim(wRsInHd!InChr) + "/" + Trim(CStr(wRsInHd!InNo))
    SetRepText Rep, "TxtInvDt", Format(wRsInHd!InPrnDt, "dd") + "/" + Format(wRsInHd!InPrnDt, "mm") + "/" + Format(wRsInHd!InPrnDt, "yyyy")
    SetRepText Rep, "TxtInvExpNo", wRsInHd!InExpNo
    GRP_REP.SetFormula Rep, "wInVaDC", wRsInHd!InVaDC
    GRP_REP.SetFormula Rep, "wInVaX", wRsInHd!InVaX
    'SetRepText Rep, "TxtCnvRt", Format(wRsInHd!InCnvRt, "#####0.00") '*** (Bef 2.13)
    SetRepText Rep, "TxtCnvRt", Format(wRsInHd!InCnvRt, "#####0.0000")  '*** (Jen 2.13)
    
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
    SetRepText Rep, "TxtGrNo", IIF(wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y", wRsInHd!InCircularNo, wRsInHd!InGrNo)
    Rep.LblGrNo.Suppress = wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y"
    Rep.LblGrDt.Suppress = wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y"
    
    If wRsInHd!InGrDt <> "01/01/80" And Not (wRsInHd!InFob < 25000 And ws_FlagLocalExp = "Y") Then _
    SetRepText Rep, "TxtGrDt", Format(wRsInHd!InGrDt, "dd") + "/" + Format(wRsInHd!InGrDt, "mm") + "/" + Format(wRsInHd!InGrDt, "yyyy")
    ' **** Manali 3.6.0 - 17/11/09 - Circular No Printed instead of GR No (InFob < 25000 and DTA )
    
    SetRepText Rep, "TxtInHValAddnFt", wRsInHd!InHValAddnFt
    SetRepText Rep, "TxtAuthSgn", wRsInHd!InhAuthSgn
    w_InHValAddMtlColYN = wRsInHd!InHValAddMtlColYN
    w_InComPer = wRsInHd!InComPer
    ws_CurCd = wRsInHd!CmCurCd    '*** (Jen 2.13)
  End If
  
  '*** Bef (2.12 Patch)
  'Dim w_LocalExp As String
  'w_LocalExp = moCn.GetFldVal("Select hExpLocal From Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  '*** Bef (2.12 Patch)
  
  If (ws_FlagLocalExp = "Y" Or adc("UOYN5") = "Y") Then     '****** Sachin 3.5.0
     Rep.CommSec.Suppress = True
  ElseIf ws_FlagLocalExp = "N" Then
     'Rep.CommSec.Suppress = False
  End If
 
'  If w_InHValAddMtlColYN = "Y" Then
''        Rep.hValOnMtl.Suppress = False
''        Rep.hValOnMtl1.Suppress = False
''        Rep.hMtlPrct.Suppress = False
''        Rep.hMtlPrct1.Suppress = False
''        Rep.hCol13.Suppress = False
''        Rep.hCol14.Suppress = False
'        REP.FldG2MtlVal.Suppress = False
'        REP.FldG2MtlPrct.Suppress = False
'        REP.FldG2TotVal.Suppress = False
''        Rep.FldG1MtlVal.Suppress = False
''        Rep.FldG1MtlPrct.Suppress = False
'  ElseIf w_InHValAddMtlColYN = "N" Then
''        Rep.hValOnMtl.Suppress = True
''        Rep.hValOnMtl1.Suppress = True
''        Rep.hMtlPrct.Suppress = True
''        Rep.hMtlPrct1.Suppress = True
''        Rep.hCol13.Suppress = True
''        Rep.hCol14.Suppress = True
'        REP.FldG2MtlVal.Suppress = True
'        REP.FldG2MtlPrct.Suppress = True
'        REP.FldG2TotVal.Suppress = True
''        Rep.FldG1MtlVal.Suppress = False
''        Rep.FldG1MtlPrct.Suppress = False
'  End If
 ' Rep.wCoCd1.SetText gs_CoNm
 
' ******* Manali - This part is shifted to DispCoNm
''  Dim ws_HName As String
''  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
''                                                      "HCd = '" + ctSelfCmCd + "' ")
''  Rep.wCoCd1.SetText ws_HName
' ******* Manali - This part is shifted to DispCoNm


  GRP_REP.SetFormula Rep, "wInComPer", w_InComPer
  If Trim(adc("UoInGrp")) <> "" Then
    GRP_REP.SetFormula Rep, "wGrp1", wGrpFormula1
    GRP_REP.SetFormula Rep, "wGrp1Desc", wGrpDescFormula1
    
    SetSubFormula Rep.SubRep2, "wSubGrp1", wGrpFormula1      ' Manali 3.8.0
  End If
  If Trim(adc("UoInGrp1")) <> "" Then
    GRP_REP.SetFormula Rep, "wGrp2", wGrpFormula2
    GRP_REP.SetFormula Rep, "wGrp2Desc", wGrpDescFormula2
   
    SetSubFormula Rep.SubRep2, "wSubGrp2", wGrpFormula2      ' Manali 3.8.0
  End If
  
  'If Trim(adc("UoInGrp")) = "" Then Rep.G1H.Suppress = True: Rep.G1F.Suppress = True: GRP_REP.SetFormula Rep, "wResSrNo", ""
  If Trim(adc("UoInGrp")) = "" Then
    GRP_REP.SetFormula Rep, "wGrp1", ""
    GRP_REP.SetFormula Rep, "wGrp1Desc", ""
    GRP_REP.SetFormula Rep, "wResSrNo", ""
  
    SetSubFormula Rep.SubRep2, "wSubGrp1", ""
  
  End If
  
  ' **** Manali 3.8.0
  If UCase(adc.MenuCd) = UCase("RepValAddnDta") Or UCase(adc.MenuCd) = UCase("RepStudStmt") Then
    Dim ws_StudDet As String, wStr6 As String
    ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
    '3.11.0
    wStr6 = " Select IdSr,qIrSrNo,IdDmCd,IdDmSz,qIrZCdRt, RmCtg, RmSCtg, " + _
            " qDCSWt as qDCSWt, 0 as qXWt,0 as qXIrRmQty ,  RmCtg+' ' +RmSCtg+ Space(5-Len(RmSCtg))+ Str(qCstRt, 9, 3) as qSCtgRt, " + _
            " qCstRt, IrRmQty as IrRmQty, qApportnWt, IdGldLs, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo, " + _
            " (Case when IrRmQty>=qCnt Then 1 Else 0 End) +  " + _
            "    Round(((IrRmQty-(Case when IrRmQty>=qCnt Then qCnt Else 0 End))*qApportnWt/(Case When qDCsWt=0 Then 1 Else qDCSWt End)),0) as Qty, " + _
            " qCnt,ROW_NUMBER() OVER (PARTITION BY IdSr,qIrZCdRt ORDER BY qIrSrNo) as qRunSr Into #TmpE " + _
            " from (Select *,(case when qFinalReq>0 and qFinalAvl> 0 then " + _
            "       (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)*(Case When (RmCtg='D' Or RmCtg='C') Then 5 Else 1 End) as qApportnWt From #TmpD) D " + _
            "               Join (Select IdSr as qIdSr,qIrZCdRt as qIrRt,Count('x') as qCnt From #TmpD D1 " + _
            "        Where D1.qIrSrNo<>0 And D1.qFinalReq>0 and D1.qFinalAvl> 0 Group By IdSr,qIrZCdRt) c " + _
            "               On qIdSr=D.IdSr And qIrRt=D.qIrZCdRt " + _
            " Where RmCtg In ('D', 'C', 'X') and (case when qFinalReq>0 and " + _
            " qFinalAvl> 0 then (case when qFinalReq>= qFinalAvl then qFinalAvl else qFinalReq end) else 0 end)>0 "
    
    moCn.CreateTmpTable "#TMPE", wStr6
    
    'Before 3.11.0
    ws_StudDet = " Select IdSr, IdDmCd, IdDmSz, RmCtg, RmSCtg, qDCSWt, 0 as qXWt, 0 as qXIrRmQty, qSCtgRt, qCstRt," + _
                 "      (Case when qRunSr=qCnt Then (Case when (IrRmQty-IsNull((Select Sum(Qty) From #TmpE a " + _
                 "      where a.qRunSr<b.qRunSr and a.IdSr=b.IdSr and a.qIrZCdRt=b.qIrZCdRt),0))>0 Then " + _
                 "      (IrRmQty-IsNull((Select Sum(Qty) From #TmpE a where a.qRunSr<b.qRunSr and a.IdSr=b.IdSr and a.qIrZCdRt=b.qIrZCdRt),0)) Else 0 End) " + _
                 "      Else Qty End) as qQty, IrRmQty, qApportnWt, IdGldLs, IdKt, IdVaCtg, DmCtg, DmSalCtg, DmUom, OdDmCol, OrRmCtg, IdTrayNo " + _
                 "   From #TmpE b  " + _
                 " Union All " + _
                 " Select max(IdSr) as IdSr, max(IdDmCd) as IdDmCd, max(IdDmSz) as IdDmSz, '' as RmCtg, '' as RmSCtg, 0 as qDCSWt, Sum(qXWt) as qXWt, Sum(qXIrRmQty) as qXIrRmQty, " + _
                 "      Space(18) as qSCtgRt, 0 as qCstRt,0 as qQty, 0 as IrRmQty,  " + _
                 "      0  as qApportnWt, " + _
                 "      max(IdGldLs) as IdGldLs, max(IdKt) as IdKt, max(IdVaCtg) as IdVaCtg, max(DmCtg) as DmCtg, max(DmSalCtg) as DmSalCtg, max(DmUom) as DmUom, max(OdDmCol) as OdDmCol, max(OrRmCtg) OrRmCtg, max(IdTrayNo) as IdTrayNo " + _
                 " from #TmpB  " + _
                 " Group By " + IIF(wDetSrt = "", ws_GrpBy, "IdSr ")

    Call moCn.RepRes(Rep.SubRep2.OpenSubreport, ws_StudDet + ctMaxDopOpt) '****** Sachin 3.01 added ctMaxDopOpt
  
    Dim wRecSelFormula As String
    If wDetSrt = "" Then
      Rep.SubRep2.OpenSubreport.RecordSelectionFormula = "{@wSubGrp1} = {?Pm-@wGrp1} and {@wSubGrp2} = {?Pm-@wGrp2} "
    ElseIf adc("UoInGrp") = "" Then
      Rep.SubRep2.OpenSubreport.RecordSelectionFormula = "{@wSubGrp3} = {?Pm-@wGrp3} "
    End If

  End If
  ' **** Manali 3.8.0
  
  
' **** Zubin 211 **** '
  With Rep

    '*** Ver 1-12
    If UCase(adc("UoPMCdFr")) = "STUDCST" Then
      GRP_REP.SetFormula Rep, "wTotMetalXCost", "{@wTotMetalCost}"
      GRP_REP.SetFormula Rep, "wTotDCXCost", "{@wDiaCsCstVal}+ {@wTotXCost}"
      'GRP_REP.SetFormula Rep, "wTotRmQty", "If {rdo.qIrSrNo}=0 Then {rdo.IrRmQty} Else 0"
      .Fldg2TotXCost1.Suppress = True
      .Fldg2TotMetalXCost.Suppress = True
      If UCase(adc.MenuCd) = UCase("RepValAddn") Then
        .Fldg2TotXCost2.Suppress = False
        .Fldg2TotDCXCost.Suppress = False
        .Fldg2TotXCost2Sub.Suppress = True
        .Fldg2TotDCXCostSub.Suppress = True
        .TxtStudDet.SetText "Details Studding/ Finding"
      Else
        .Fldg2TotXCost2.Suppress = True
        .Fldg2TotDCXCost.Suppress = True
        .G2HX.Suppress = True
        .Fldg1XWt.Suppress = True
        .FldgXWt.Suppress = True
        GRP_REP.SetFormula Rep, "g1_ApportnWtCts", "Sum ({@wDCXApportnWt}, {@wGrp1})"
        GRP_REP.SetFormula Rep, "g_ApportnWtCts", "Sum ({@wDCXApportnWt})"
        .TxtStudDetSub.SetText "Details Studding/ Finding"
      End If
    ElseIf UCase(adc("UoPMCdFr")) = "METCST" Then
      GRP_REP.SetFormula Rep, "wTotMetalXCost", "{@wTotMetalCost}+ {@wTotXCost}"
      GRP_REP.SetFormula Rep, "wTotDCXCost", "{@wDiaCsCstVal}"
      'GRP_REP.SetFormula Rep, "wTotRmQty", "If {rdo.qIrSrNo}= 0 and (Trim({rdo.RmCtg})= 'D' or Trim({rdo.RmCtg})= 'C') Then {rdo.IrRmQty} Else 0"
      .Fldg2TotXCost1.Suppress = False
      .Fldg2TotMetalXCost.Suppress = False
      .Fldg2TotXCost2.Suppress = True
      .Fldg2TotDCXCost.Suppress = True
        .Fldg2TotXCost2Sub.Suppress = True
        .Fldg2TotDCXCostSub.Suppress = True
      .TxtStudDet.SetText "Details Of Studding"
      .TxtStudDetSub.SetText "Details Of Studding"
      .SubRep2_G2F.Suppress = True
    End If
    '*** Ver 1-12
    
    '****   urmi 2.04-3 Show Running Sr or Not
    '''''  If adc("UoYN") = "Y" Then
    '''Geeta***Or adc("UoInDet") = ""
    If adc("UoInDet") = "INVSR/SR" Or adc("UoInDet") = "DMCD/SR" Or adc("UoInDet") = "" Then
    ' **** Zubin 211 **** '
      .FldRunningSrNo.Suppress = False
      .FldIdSr.Suppress = True
    Else
      .FldRunningSrNo.Suppress = True
      .FldIdSr.Suppress = False
    End If
    '******************************************
    ' **** Zubin 211 **** '
    
    If UCase(adc("UoYN1")) = "Y" Then
      .TxtAnnexAtt.Suppress = False
    ElseIf UCase(adc("UoYN1")) = "N" Then
      .TxtAnnexAtt.Suppress = True
    End If
    
    '*** (Jen 2.13)
    .hTotValStud.SetText "Total Value" + Chr(13) + "in " + ws_CurCd
    .hFOBInUSD.SetText "FOB Value" + Chr(13) + "in " + ws_CurCd
    .hTotInUSD.SetText "Total" + Chr(13) + "in " + ws_CurCd
    '*** (Jen 2.13)
  
    If UCase(adc.MenuCd) = UCase("RepValAddn") Then
      '*** Jen (2.12 Patch)
      If ws_FlagLocalExp = "N" And adc("UOYN5") = "N" Then        '****** Sachin 3.5.0
      '*** Jen (2.12 Patch)
        
        '*** (Jen 2.13)
        '.hValOnMtl.SetText "On Metal" + Chr(13) + "Val in US$"
        .hValOnMtl.SetText "On Metal" + Chr(13) + "Val in " + ws_CurCd
        .hValOnMtl1.SetText "12 - 15"
              
        GRP_REP.SetFormula Rep, "g2_MetalVal", "Sum ({@wMetalVal}, {@wGrp2})"
        GRP_REP.SetFormula Rep, "g1_MetalVal", "Sum ({@wMetalVal}, {@wGrp1})"
        GRP_REP.SetFormula Rep, "g_MetalVal", "Sum ({@wMetalVal})"
        '*** (Jen 2.13)
        
        '*** (Jen 2.14 Next Patch 1)
        If UCase(adc("UOYN4")) = "Y" Then
          .FldG2TotVal.Suppress = True: .FldG1TotVal.Suppress = True: .FldGTotVal.Suppress = True
          .hTotInUSD.Suppress = True: .hTotInUSD1.Suppress = True: .hCol12.Suppress = True
        Else
          If w_InHValAddMtlColYN = "N" Then
            .FldG2TotVal.Suppress = True
          Else
            .FldG2TotVal.Suppress = False
            .FldG1TotVal.Suppress = False
            .FldGTotVal.Suppress = False
          End If
        End If
        '*** (Jen 2.14 Next Patch 1)
        
        If UCase(adc("UOYN2")) = "Y" Then
          '.FldG2TotVal.Suppress = True: .FldG1TotVal.Suppress = True: .FldGTotVal.Suppress = True   '*** (Bef 2.14 Next)
          .FldG2MtlVal.Suppress = True: .FldG2MtlPrct.Suppress = True
          .FldG2StudVal.Suppress = True: .FldG2XVal.Suppress = True: .FldG2ValDC.Suppress = True
          .FldG2ValX.Suppress = True: .FldG1MtlVal.Suppress = True
          .FldG1MtlPrct.Suppress = True: .FldG1StudXVal.Suppress = True
          .FldGMtlVal.Suppress = True: .FldGMtlPrct.Suppress = True: .FldGStudXVal.Suppress = True
          
          '*** 30/09/05
          '.hTotInUSD.Suppress = True: .hTotInUSD1.Suppress = True: .hCol12.Suppress = True   '*** (Bef 2.14 Next)
          .hValOnMtl.Suppress = True: .hValOnMtl1.Suppress = True: .hCol13.Suppress = True
          .hMtlPrct.Suppress = True: .hMtlPrct1.Suppress = True: .hCol14.Suppress = True
          .hStud.Suppress = True: .hCol15.Suppress = True
          .hStudPrc.Suppress = True: .hCol16.Suppress = True
          .hDetsOfValAddn.Suppress = True
          '*** 30/09/05
        Else
          If w_InHValAddMtlColYN = "N" Then
            .FldG2MtlVal.Suppress = True
            .FldG2MtlPrct.Suppress = True
            '.FldG2TotVal.Suppress = True     '*** (Bef 2.14 Next Patch 1)
          Else
            '.FldG2TotVal.Suppress = False     '*** (Bef 2.14 Next Patch 1)
            .FldG2MtlVal.Suppress = False: .FldG2MtlPrct.Suppress = False
            .FldG2StudVal.Suppress = False: .FldG2XVal.Suppress = False: .FldG2ValDC.Suppress = False
            .FldG2ValX.Suppress = False
            '.FldG1TotVal.Suppress = False    '*** (Bef 2.14 Next Patch 1)
            .FldG1MtlVal.Suppress = False
            .FldG1MtlPrct.Suppress = False: .FldG1StudXVal.Suppress = False
            '.FldGTotVal.Suppress = False     '*** (Bef 2.14 Next Patch 1)
            .FldGMtlVal.Suppress = False: .FldGMtlPrct.Suppress = False: .FldGStudXVal.Suppress = False
          End If
        End If
        
      '*** (Jen 2.13)
      Else
      
        .hValOnMtl.SetText "% On Total" + Chr(13) + "Cost"
        
        If adc("UOYN5") = "N" Then
          .hValOnMtl1.SetText "12/ (9 + 10)"
          GRP_REP.SetFormula Rep, "g2_MetalVal", "if ({@g2_TotMetalXCost}+ {@g2_TotDCXCost})> 0 then " + _
                                                 "    {@g2_TotVal}/ ({@g2_TotMetalXCost}+ {@g2_TotDCXCost})* 100 " + _
                                                 "Else 100"
          GRP_REP.SetFormula Rep, "g1_MetalVal", "if ({@g1_TotMetalXCost}+ {@g1_TotDCXCost})> 0 then " + _
                                                 "    {@g1_TotVal}/ ({@g1_TotMetalXCost}+ {@g1_TotDCXCost})* 100 " + _
                                                 "Else 100"
          GRP_REP.SetFormula Rep, "g_MetalVal", "if ({@g_TotMetalXCost}+ {@g_TotDCXCost})> 0 then " + _
                                                 "    {@g_TotVal}/ ({@g_TotMetalXCost}+ {@g_TotDCXCost})* 100 " + _
                                                 "Else 100"
        Else
          '****** Sachin 3.5.0
          .hValOnMtl1.SetText "12 / 11"
          GRP_REP.SetFormula Rep, "g2_MetalVal", " if ({@g2_fob})> 0 Then {@g2_TotVal}/{@g2_fob}*100 Else 0 "
          GRP_REP.SetFormula Rep, "g1_MetalVal", " if ({@g1_fob})> 0 Then {@g1_TotVal}/{@g1_fob}*100  Else 0 "
          GRP_REP.SetFormula Rep, "g_MetalVal", " if ({@g_fob})> 0 Then {@g_TotVal}/{@g_fob}*100  Else 0 "
          '****** Sachin 3.5.0
        End If
                                                       
        ' ***** Manali 3.6.0 - Hide Column 12 and 13 for DTA
        If UCase(adc("UOYN4")) = "Y" Then
          .FldG2TotVal.Suppress = True: .FldG1TotVal.Suppress = True: .FldGTotVal.Suppress = True
          .hTotInUSD.Suppress = True: .hTotInUSD1.Suppress = True: .hCol12.Suppress = True
        Else
          .FldG2TotVal.Suppress = False: .FldG1TotVal.Suppress = False: .FldGTotVal.Suppress = False
          .hTotInUSD.Suppress = False: .hTotInUSD1.Suppress = False: .hCol12.Suppress = False
        End If
                                                 
        If UCase(adc("UOYN2")) = "Y" Then
          .FldG2MtlVal.Suppress = True: .FldG1MtlVal.Suppress = True:
          .FldGMtlVal.Suppress = True
          .hValOnMtl.Suppress = True: .hValOnMtl1.Suppress = True: .hCol13.Suppress = True
          .hDetsOfValAddn.Suppress = True
        Else
          .FldG2MtlVal.Suppress = False
          .FldG1MtlVal.Suppress = False
          .FldGMtlVal.Suppress = False
          .hValOnMtl.Suppress = False: .hValOnMtl1.Suppress = False: .hCol13.Suppress = False
          .hDetsOfValAddn.Suppress = False
        End If
        ' ***** Manali 3.6.0 - Hide Column 12 and 13 for DTA
                                                 
        .hMtlPrct.Suppress = True: .hMtlPrct1.Suppress = True
        .hCol14.Suppress = True: .FldG2MtlPrct.Suppress = True
        .FldG1MtlPrct.Suppress = True: .FldGMtlPrct.Suppress = True
        .LinehCol14.Suppress = True: .LineDetCol14.Suppress = True
        
        .hStud.Suppress = True: .hCol15.Suppress = True
        .FldG2StudVal.Suppress = True: .FldG2XVal.Suppress = True
        .FldG1StudXVal.Suppress = True: .FldGStudXVal.Suppress = True
        .LinehCol15.Suppress = True: .LineDetCol15.Suppress = True
        
        .hStudPrc.Suppress = True: .hCol16.Suppress = True
        .FldG2ValDC.Suppress = True: .FldG2ValX.Suppress = True
      '*** (Jen 2.13)
      End If
      
      ' ***** Manali 3.8.0
      GRP_REP.SetFormula Rep, "wDtaRep", "'N'"
      .PHDetDta.Suppress = True:
      .LineSubColFobVal.Suppress = True: .LineSubColShp.Suppress = True
      '.G2HStudSubrep.Suppress = True:
      .FldG2FobSub.Suppress = True
      .FldG1FobSub.Suppress = True: .Fldg1DiaCsCstValSub.Suppress = True
      .FldG1ApprtnWtCtsSub.Suppress = True: .FldG1RmQtySub.Suppress = True
      .FldGFobSub.Suppress = True: .FldGRmQtySub.Suppress = True
      .FldGApprtnWtCtsSub.Suppress = True: .FldgDiaCsCstValSub.Suppress = True
      .FldGTotDcxCstSub.Suppress = True: .FldG1TotDcxCstSub.Suppress = True
      ' ***** Manali 3.10.0 - 14/07/12 - Total Value in DTA report
      .FldgTotValSub.Suppress = True: .Fldg1TotValSub.Suppress = True: .Fldg2TotValSub.Suppress = True
    Else
      GRP_REP.SetFormula Rep, "wDtaRep", "'Y'"
      .PHDET.Suppress = True
      .FldG2RmQty.Suppress = True: .FldG2ApprtnWtCts.Suppress = True
      .FldG2DiaCsCstVal.Suppress = True: .FldG2FobDet.Suppress = True
      .FldG2TotVal.Suppress = True: .FldG2MtlVal.Suppress = True
      .FldG2MtlPrct.Suppress = True: .FldG2StudVal.Suppress = True
      .FldG2ValDC.Suppress = True: .FldG2XVal.Suppress = True
      .FldG2ValX.Suppress = True
      .FldG1RmQtyDet.Suppress = True: .FldG1ApprtnWtCtsDet.Suppress = True
      .FldG1TotDcxCstDet.Suppress = True: .FldG1FobDet.Suppress = True
      .FldG1TotVal.Suppress = True: .FldG1MtlVal.Suppress = True
      .FldG1MtlPrct.Suppress = True: .FldG1StudXVal.Suppress = True
      
      .FldGRmQtyDet.Suppress = True: .FldGApprtnWtCtsDet.Suppress = True
      .FldGTotDcxCstDet.Suppress = True: .FldGFobDet.Suppress = True
      .FldGTotVal.Suppress = True: .FldGMtlVal.Suppress = True
      .FldGMtlPrct.Suppress = True: .FldGStudXVal.Suppress = True
      
      .LineDetColPcs.Suppress = True: .LineDetColMetVal.Suppress = True
      .LineDetCol14.Suppress = True: .LineDetCol15.Suppress = True
'      .LineDetCol16.Suppress = True
'      .LineG1HExt1.Suppress = True: .LineG1FExt1.Suppress = True
'      .LineG1FExt2.Suppress = True: .LineRFExt1.Suppress = True
'      .LineRFExt2.Suppress = True

        If UCase(adc.MenuCd) = UCase("RepStudStmt") Then  '3.11.1
            If UCase(adc("UOYN4")) = "Y" Then       'Hide Val Addn Dets
               .Fldg2TotValSub.Suppress = True: .Fldg1TotValSub.Suppress = True
               .hDetOfValAdd.Suppress = True: .FldgTotValSub.Suppress = True: .hTotUSD.Suppress = True
            Else
              .Fldg2TotValSub.Suppress = False: .hDetOfValAdd.Suppress = False
              .Fldg1TotValSub.Suppress = False: .FldgTotValSub.Suppress = False: .hTotUSD.Suppress = False
            End If
            If UCase(adc("UOYN2")) = "Y" Then       'Hide FOB Dets
              .FldG2FobSub.Suppress = True
            Else
              .FldG2FobSub.Suppress = False
            End If
        End If
        
    End If
    
    '3.11.1.0
    Rep.PHDetStudStmt.Suppress = True
    If UCase(adc.MenuCd) = UCase("RepStudStmt") Then
        .PHDetStudStmt.Suppress = False
        .PHDetDta.Suppress = True
        .g1TotMetWt.Suppress = True
        .g1TotMetPureWt.Suppress = False
        .gTotMetWt.Suppress = True
        .gTotMetPureWt.Suppress = False
    End If
    
    ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
    If ws_FlagLocalExp = "Y" And adc("UoYn6") = "N" Then
      Rep.SecXBoe.Suppress = True
    Else
      Rep.SecXBoe.Suppress = False
    End If
    ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
    
    '3.11.0
    If adc("UoYn7") = "N" Then
      Rep.SubRepXRmKt.Suppress = True
    Else
      Rep.SubRepXRmKt.Suppress = False
    End If
    
    '*** (Jen 2.13)
    'If adc("UoYN3") = "Y" Then
    If Trim(UCase(moCn.GetFldVal("Select min(InShowGldLs) from InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd= '" + adc("UoCoCdFr") + "' and " + _
                                 "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
                                 "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " "))) = "Y" Then
      ''.TxtLossPrc.Suppress = False: .g2GLossPrc.Suppress = False
      ''.g2PLossPrc.Suppress = False: .g2SLossPrc.Suppress = False
    Else
      .TxtLossPrc.Suppress = True: .g2GLossPrc.Suppress = True
      .g2PLossPrc.Suppress = True: .g2SLossPrc.Suppress = True
      .g2LLossPrc.Suppress = True   ' ***** Manali 3.03 - 30/06/08 - Pd Mod - 'L'
    End If
    '*** (Jen 2.13)
  
  End With
' **** Zubin 211 **** '
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
'*** (Jen 2.12)
  Exit Sub

RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
'*** (Jen 2.12)
End Sub
Private Sub ApportnTbl()
  Dim wStr1 As String, wStr2 As String, wStr3 As String
  '*** Creating Temporary Table #TMPD which is the main Apportioning Sql ******
  
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - PdWt added
  ' ***** Manali 3.8.0 - IdTrayNo added in Sort, xIrRmQty added (ValAddnDTA)
  wStr1 = "Create table #TMPD " + _
          "(  qIrSrNo     integer         null, " + _
          "   IdSr        integer         null, " + _
          "   IdDmCd      varchar(15)     null, " + _
          "   IdDmSz      varchar(5)      null, " + _
          "   IdQty       float           null, " + _
          "   IdGldLs     float           null, " + _
          "   qIrWt       float           null, " + _
          "   qGrsWt      float           null, " + _
          "   qNetWt      float           null, " + _
          "   qGldWt      float           null, " + _
          "   qPlWt       float           null, " + _
          "   qSlWt       float           null, " + _
          "   qPdWt       float           null, " + _
          "   qMWt        float           null, " + _
          "   qNetPureWt      float           null, " + _
          "   qGldNetWt      float           null, " + _
          "   qPlNetWt       float           null, " + _
          "   qSlNetWt       float           null, " + _
          "   qPdNetWt       float           null, " + _
          "   qMNetWt        float           null, "

  
  '****** (Jen 2.13) added fields qGldLsWt, qPlLsWt, qSlLsWt ******'
  ' ***** Manali 3.03 - 30/06/08 - Pd Mod - PdLsWt added
  wStr2 = "   qGldLsWt    float           null, " + _
          "   qPlLsWt     float           null, " + _
          "   qSlLsWt     float           null, " + _
          "   qPdLsWt     float           null, " + _
          "   qMLsWt     float           null, " + _
          "   qXWt        float           null, " + _
          "   qDCSWt      float           null, " + _
          "   qXIrRmQty   float           null, " + _
          "   RmCtg       varchar(2)      null, " + _
          "   RmSCtg      varchar(5)      null, " + _
          "   IrRmZCd     varchar(12)     null, " + _
          "   qTyp        varchar(50)     null, " + _
          "   IrRmQty     integer         null, " + _
          "   qSalRt      float           null, " + _
          "   qIrZCdRt    varchar(21)     null, " + _
          "   qDCSSalVal  float           null, "
  
          '*** (Bef 2.14)
          '"   qTyp        varchar(30)     null, "
          '*** (Bef 2.14)
          
          '*** (Jen 2.14)
          '"   qTyp        varchar(50)     null, "
          '*** (Jen 2.14)

  ' Zubin 212 (DmSalCtg added)
  wStr3 = "   qAvgSalPrc  float           null, " + _
          "   qFOB        float           null, " + _
          "   IdKt        varchar(5)      null, " + _
          "   IdVaCtg     varchar(5)      null, " + _
          "   DmCtg       varchar(5)      null, " + _
          "   DmSalCtg    varchar(8)      null, " + _
          "   DmUom       varchar(5)      null, " + _
          "   OdDmCol     varchar(5)      null, " + _
          "   OrRmCtg     varchar(2)      null, " + _
          "   IdTrayNo    integer         null, " + _
          "   TBOENo      varchar(20)     null, " + _
          "   TBOEDt      smalldatetime   null, " + _
          "   qCstRt      float           null, " + _
          "   qFinalReq   float           null, " + _
          "   qFinalAvl   float           null " + _
          ")"
   moCn.CreateTmpTable "#TMPD", wStr1 + wStr2 + wStr3

End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowRunSr"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "Annex"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  Case Is = 2     'Tag = "HidCol13_16"
    If .Value = Checked Then adc("UoYN2") = "Y" Else adc("UoYN2") = "N"
  Case Is = 4     'Tag = "HideCol12"
    If .Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"
  Case Is = 5     'Tag = "ShowProfit"
    If .Value = Checked Then adc("UoYN5") = "Y" Else adc("UoYN5") = "N"
  ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
  Case Is = 6     'Tag = "ShowXBoe"
    If .Value = Checked Then adc("UoYN6") = "Y" Else adc("UoYN6") = "N"
  ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
  
  Case Is = 7     'Tag = ""     3.11.0
    If .Value = Checked Then adc("UoYN7") = "Y" Else adc("UoYN7") = "N"
  
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
    Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gs_PTyp = "FNDOPT"
    gs_Tbl = "Param"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    '*** Jen (2.12 Patch)
    ws_FlagLocalExp = moCn.GetFldVal("Select hExpLocal from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                                  "HCd= '" + ctSelfCmCd + "' ")
    ' ***** Manali 3.6.0 - 01/09/09
''''    If ws_FlagLocalExp = "Y" Then
''''      LblHideCols.Visible = False
''''      adc("UOYN2").Visible = False
''''      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' ****** Manali 3.5.0 - 05/12/08 - Check Boxes added
''''
''''      '*** (Jen 2.14 Next Patch 1)
''''      LblHideCol12.Visible = False
''''      adc("UOYN4").Visible = False
''''      '*** (Jen 2.14 Next Patch 1)
''''      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' ****** Manali 3.5.0 - 05/12/08 - Check Boxes added
''''    End If
    '*** Jen (2.12 Patch)
    
    ' **** Zubin 211 **** '
    LblRunnSr.Visible = False
    adc("UOYN").Visible = False
    ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' ****** Manali 3.5.0 - 05/12/08 - Check Boxes added
    ' **** Zubin 211 **** '
    
    ' ***** Manali 3.8.0
    If UCase(adc.MenuCd) <> UCase("RepValAddn") Then
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False
        If UCase(adc.MenuCd) = UCase("RepStudStmt") Then
          ChkBoxArr(2).Enabled = True: ChkBoxArr(2).Visible = True
          ChkBoxArr(4).Enabled = True: ChkBoxArr(4).Visible = True
          ChkBoxArr(2).Caption = "Hide FOB Dets": ChkBoxArr(4).Caption = "Hide Value Addn"
        End If
    End If
    ' ***** Manali 3.8.0
        
    Call SetGroupSort
    
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
    gs_PTyp = "FNDOPT"
    gs_Tbl = "Param"
  
 'urmi 2.04-3
  Case Is = UCase("UoInGrp1")
    If pv_NewValue <> "" Then adc("UoYn") = ""
  
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoPMCdFr")
    gs_PTyp = "FNDOPT"
    gs_Tbl = "Param"
 'urmi 2.04-3
  Case Is = UCase("UoYn")
     If adc("UoInGrp1") <> "" Then Cancel = True: DispMsg "Cannot edit when 2nd Grp opt is not blank", etWarning
  
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
 
  '*** (2.12)
  'If adc("UoInGrp1") = "" And adc("UoInDet") = "" Then adc("UoInDet") = "INVSR"
  '*** (2.12)
  
' **** Zubin 211 **** '
  'If adc("UoInDet") = "" Then adc("UoInDet") = "INVSR"     '(Bef 2.12)
  If adc("UoYN1") = "" Then adc("UoYN1") = "Y"
  If adc("UoYN2") = "" Then adc("UoYN2") = "N"
  If adc("UoYN4") = "" Then adc("UoYN4") = "N"      '*** (Jen 2.14 Next Patch 1)
' **** Zubin 211 **** '

  '****** Sachin 3.5.0 ******
  adc("UOYN5") = "N"
  ' ***** Manali 3.8.0
  If UCase(adc.MenuCd) = UCase("RepValAddn") Then
    If ws_FlagLocalExp = "Y" Then
      adc("UOYN5").Visible = False: LblProfitPer.Visible = False
      ChkBoxArr(5).Enabled = False: ChkBoxArr(5).Visible = False      ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
      ChkBoxArr(6).Enabled = True: ChkBoxArr(6).Visible = True      ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
    Else
      'adc("UOYN5").Visible = LblProfitPer.Visible = True: LblProfitPer.Visible = True
      ChkBoxArr(5).Enabled = True: ChkBoxArr(5).Visible = True        ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
    End If
  '****** Sachin 3.5.0 ******
  
  '3.11.1.0
  ElseIf UCase(adc.MenuCd) = UCase("RepStudStmt") Then
        adc("UoPMCdFr") = "METCST"
        adc("UoPMCdFr").Visible = False
        LblFndCst.Visible = False
  End If
  '3.11.1.0
  
  adc("UoYN3") = "Y"      '*** (Jen 2.13)
  
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN6") = "" Then adc("UoYn6") = "Y"        ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
  If adc("UoYN7") = "" Then adc("UoYn7") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  If adc("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If adc("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If adc("UoYN5") = "Y" Then ChkBoxArr(5).Value = Checked Else ChkBoxArr(5).Value = Unchecked
  If adc("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked   ' **** Manali 3.6.0 - 17/11/09 - Show Findings Boe for DTA
  If adc("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked   ' **** 3.11.0
  ' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
 
End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
    Dim ws_HName As String
    ' ***** Manali - CoName from CustMst instead of head
    ws_HName = moCn.GetFldVal("Select CmName from CustMst where CmCtg='Z' and CmCd='" + adc("UoCoCdFr") + "' ")
    'ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    Rep.wCoCd.SetText ws_HName
    Rep.wCoCd1.SetText "For  " + ws_HName
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
    Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    'Set mRep = Nothing   '*** (Bef 2.13)
    Set Rep = Nothing
    'Set Rep1 = Nothing   '*** (Bef 2.13)
    '*** (09/08/05)
End Sub


