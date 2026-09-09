VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpFgCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Printing"
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
   Begin MwfCtl.MWCTL_BTN1 CmdDispSc 
      Height          =   390
      Left            =   3210
      TabIndex        =   131
      ToolTipText     =   "Display Scope Options"
      Top             =   9660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Disp&lay "
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   1890
      TabIndex        =   62
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9660
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Extra Sc&ope"
      BeginProperty LblFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10710
      TabIndex        =   64
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
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   9600
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4680
      TabIndex        =   73
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   74
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
         TabIndex        =   75
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
      Left            =   5820
      TabIndex        =   72
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   69
         Top             =   -135
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
      TabIndex        =   65
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
      Height          =   9660
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   0
         TabIndex        =   67
         Top             =   120
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFgCat.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFgCat.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   71
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   66
               Top             =   360
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8835
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   14865
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1920
               TabIndex        =   167
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   6000
               Width           =   3825
               Begin VB.OptionButton OptBag 
                  Caption         =   "All"
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
                  Left            =   2700
                  TabIndex        =   170
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Closed"
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
                  Left            =   1260
                  TabIndex        =   169
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptBag 
                  Caption         =   "Open"
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
                  TabIndex        =   168
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Distinct Dsg"
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
               Index           =   28
               Left            =   75
               TabIndex        =   29
               Tag             =   "ShowDistDsg"
               ToolTipText     =   "Check To Show Distinct Design Scope"
               Top             =   4815
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Run Sr         "
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
               Index           =   25
               Left            =   75
               TabIndex        =   31
               Tag             =   "ShowRunSr"
               ToolTipText     =   "Check To Show Voucher Serials"
               Top             =   5640
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Vch Sr         "
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
               Index           =   24
               Left            =   75
               TabIndex        =   30
               Tag             =   "ShowVchSr"
               ToolTipText     =   "Check To Show Voucher Serials"
               Top             =   5280
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Grp Comb Desc "
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
               Index           =   13
               Left            =   4590
               TabIndex        =   17
               Tag             =   "ShowGrpDesc"
               ToolTipText     =   "Check To Show Group Description"
               Top             =   1965
               Width           =   2415
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Final Sort On Dsg Cd   "
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
               Index           =   11
               Left            =   4590
               TabIndex        =   19
               Tag             =   "FinalSrt"
               ToolTipText     =   "Check To Show Picture"
               Top             =   2250
               Width           =   2415
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1905
               TabIndex        =   1
               ToolTipText     =   "Enter Tc Type"
               Top             =   825
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCTYPFR"
               ReCalcParent    =   "UOCOCDFR"
               IdName          =   "UOTCTYPFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5385
               TabIndex        =   6
               ToolTipText     =   "Enter To Order Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOFGYYTO"
               IdName          =   "UOFGYYTO"
               CmpStr          =   "FdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2685
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Year"
               Top             =   1110
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOFGYYFR"
               IdName          =   "UOFGYYFR"
               CmpStr          =   "FdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1905
               TabIndex        =   2
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGTCFR"
               IdName          =   "UOFGTCFR"
               CmpStr          =   "FdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3105
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Character"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGCHRFR"
               IdName          =   "UOFGCHRFR"
               CmpStr          =   "FdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4590
               TabIndex        =   5
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGTCTO"
               IdName          =   "UOFGTCTO"
               CmpStr          =   "FdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5805
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Character"
               Top             =   1110
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFGCHRTO"
               IdName          =   "UOFGCHRTO"
               CmpStr          =   "FdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5520
               TabIndex        =   11
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1395
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDSRTO"
               IdName          =   "UOFDSRTO"
               CmpStr          =   "FdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   2805
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1395
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDSRFR"
               IdName          =   "UOFDSRFR"
               CmpStr          =   "FdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   1905
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Number"
               Top             =   1395
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFGNOFR"
               IdName          =   "UOFGNOFR"
               CmpStr          =   "FdNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4590
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Number"
               Top             =   1395
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFGNOTO"
               IdName          =   "UOFGNOTO"
               CmpStr          =   "FdNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1905
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               ReCalcOn        =   "UOTCTYPFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "FdCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4590
               TabIndex        =   24
               ToolTipText     =   "Enter To Design Code"
               Top             =   2820
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "FdDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7290
               TabIndex        =   25
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2820
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "FdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1905
               TabIndex        =   23
               ToolTipText     =   "Enter From Design Code"
               Top             =   2820
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "FdDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4590
               TabIndex        =   27
               ToolTipText     =   "Enter To Design Size"
               Top             =   3105
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "FdDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1905
               TabIndex        =   26
               ToolTipText     =   "Enter From Design Size"
               Top             =   3105
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "FdDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4590
               TabIndex        =   28
               ToolTipText     =   "Enter To Suffix"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "FdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4590
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Category"
               Top             =   2535
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
               Index           =   30
               Left            =   1905
               TabIndex        =   20
               ToolTipText     =   "Enter From Design Category"
               Top             =   2535
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
               Index           =   8
               Left            =   7290
               TabIndex        =   22
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2535
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1905
               TabIndex        =   14
               ToolTipText     =   "Enter First Group"
               Top             =   1965
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
               Index           =   27
               Left            =   1905
               TabIndex        =   15
               ToolTipText     =   "Enter Second Group"
               Top             =   2250
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
               Index           =   42
               Left            =   7335
               TabIndex        =   18
               ToolTipText     =   "Final Groping on Design Code (Y/ N)"
               Top             =   2250
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN11"
               ReCalcOn        =   "UOYN3"
               IdName          =   "UOYN11"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7335
               TabIndex        =   16
               ToolTipText     =   "Show Picture (Yes / No)"
               Top             =   1965
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN13"
               IdName          =   "UOYN13"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   62
               Left            =   2280
               TabIndex        =   41
               ToolTipText     =   "Show Components Details (Yes / No)"
               Top             =   5265
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN24"
               IdName          =   "UOYN24"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   71
               Left            =   4590
               TabIndex        =   13
               ToolTipText     =   "Enter To Order Date"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "FdDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   1905
               TabIndex        =   12
               ToolTipText     =   "Enter From Order Date"
               Top             =   1680
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "FdDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2280
               TabIndex        =   148
               ToolTipText     =   "Show Distinct Design(Y/N)"
               Top             =   4890
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN28"
               IdName          =   "UOYN28"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   5385
               TabIndex        =   35
               ToolTipText     =   "Enter To Transaction Year"
               Top             =   3675
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   2790
               TabIndex        =   32
               ToolTipText     =   "Enter From Transaction Year"
               Top             =   3675
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   3300
               TabIndex        =   33
               ToolTipText     =   "Enter From Transaction Character"
               Top             =   3675
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   4590
               TabIndex        =   34
               ToolTipText     =   "Enter To Transaction Tc"
               Top             =   3675
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   5805
               TabIndex        =   36
               ToolTipText     =   "Enter To Transaction Character"
               Top             =   3675
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   5505
               TabIndex        =   40
               ToolTipText     =   "Enter To Transaction Serial"
               Top             =   3960
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRTO"
               IdName          =   "UOTDSRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   63
               Left            =   2805
               TabIndex        =   38
               ToolTipText     =   "Enter From Transaction Serial"
               Top             =   3960
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTDSRFR"
               IdName          =   "UOTDSRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   69
               Left            =   1905
               TabIndex        =   37
               ToolTipText     =   "Enter From Transaction Number"
               Top             =   3960
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOFR"
               IdName          =   "UOTNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   70
               Left            =   4590
               TabIndex        =   39
               ToolTipText     =   "Enter To Transaction Number"
               Top             =   3960
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTNOTO"
               IdName          =   "UOTNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   64
               Left            =   2280
               TabIndex        =   157
               ToolTipText     =   "Show Running Sr (Yes / No)"
               Top             =   5580
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN25"
               IdName          =   "UOYN25"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   1905
               TabIndex        =   158
               ToolTipText     =   "Enter From Suffix"
               Top             =   3390
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "FdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1905
               TabIndex        =   159
               ToolTipText     =   "Enter From Transaction Tc"
               Top             =   3675
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4590
               TabIndex        =   160
               ToolTipText     =   "Enter To FgSubLoc Location"
               Top             =   4245
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCTO"
               IdName          =   "UOFGSUBLOCTO"
               CmpStr          =   "BFgSubLoc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   7290
               TabIndex        =   161
               ToolTipText     =   "Enter FgSubLoc Location Selection"
               Top             =   4245
               Width           =   7530
               _ExtentX        =   13282
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOFGSUBLOCSEL"
               IdName          =   "UOFGSUBLOCSEL"
               CmpStr          =   "BFgSubLoc In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1905
               TabIndex        =   162
               ToolTipText     =   "Enter From FgSubLoc Location"
               Top             =   4245
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOFGSUBLOCFR"
               IdName          =   "UOFGSUBLOCFR"
               CmpStr          =   "BFgSubLoc>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   1905
               TabIndex        =   164
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   4530
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
               Index           =   13
               Left            =   5895
               TabIndex        =   166
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   6000
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Open Bags"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   42
               Top             =   6000
               Width           =   1185
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
               Left            =   75
               TabIndex        =   165
               Top             =   4530
               Width           =   1005
            End
            Begin VB.Label lblFgSubLoc 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Sub Loc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   75
               TabIndex        =   163
               Top             =   4245
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
               Index           =   35
               Left            =   2745
               TabIndex        =   156
               Top             =   3960
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
               Index           =   34
               Left            =   5445
               TabIndex        =   155
               Top             =   3960
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
               Height          =   375
               Index           =   33
               Left            =   2655
               TabIndex        =   154
               Top             =   3675
               Width           =   135
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
               Index           =   32
               Left            =   5325
               TabIndex        =   153
               Top             =   3675
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
               Height          =   375
               Index           =   29
               Left            =   3165
               TabIndex        =   152
               Top             =   3675
               Width           =   135
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
               Index           =   4
               Left            =   5745
               TabIndex        =   151
               Top             =   3675
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fm No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   150
               Top             =   3960
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fm Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   149
               Top             =   3720
               Width           =   1725
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   145
               Top             =   1680
               Width           =   1425
            End
            Begin VB.Label LblShowGrpDesc 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Group Comb Desc"
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
               Left            =   4590
               TabIndex        =   104
               ToolTipText     =   "Location"
               Top             =   1965
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Final Sort on Dsg Cd"
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
               Left            =   4590
               TabIndex        =   103
               ToolTipText     =   "Location"
               Top             =   2250
               Visible         =   0   'False
               Width           =   2145
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
               Left            =   75
               TabIndex        =   99
               ToolTipText     =   "Location"
               Top             =   1965
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
               Left            =   75
               TabIndex        =   98
               ToolTipText     =   "Location"
               Top             =   2250
               Width           =   1875
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
               Left            =   75
               TabIndex        =   92
               Top             =   2535
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Selection22"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   7290
               TabIndex        =   91
               Top             =   120
               Width           =   1575
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
               Left            =   75
               TabIndex        =   90
               Top             =   2820
               Width           =   1845
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
               Index           =   0
               Left            =   75
               TabIndex        =   89
               Top             =   3105
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Suffix"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   88
               Top             =   3390
               Width           =   1845
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
               Left            =   75
               TabIndex        =   87
               Top             =   540
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg Tc/Yy/Chr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   86
               Top             =   1110
               Width           =   1845
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Fg No/Sr"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   75
               TabIndex        =   85
               Top             =   1395
               Width           =   1845
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
               Left            =   5745
               TabIndex        =   84
               Top             =   1110
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
               Left            =   3045
               TabIndex        =   83
               Top             =   1110
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
               Left            =   5325
               TabIndex        =   82
               Top             =   1110
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
               Left            =   2625
               TabIndex        =   81
               Top             =   1110
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
               Left            =   5445
               TabIndex        =   80
               Top             =   1395
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
               Left            =   2745
               TabIndex        =   79
               Top             =   1395
               Width           =   105
            End
            Begin VB.Label lblTcTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   75
               TabIndex        =   78
               ToolTipText     =   "Location"
               Top             =   825
               Width           =   1845
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
               Left            =   4290
               TabIndex        =   77
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
               Left            =   1905
               TabIndex        =   76
               Top             =   120
               Width           =   1335
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   3915
         Left            =   5640
         TabIndex        =   93
         Top             =   5640
         Visible         =   0   'False
         Width           =   3885
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Design Remarks                    "
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
            Index           =   26
            Left            =   90
            TabIndex        =   140
            Tag             =   "ShowDsgRem"
            ToolTipText     =   "Check To Show Order Design Level Remark"
            Top             =   1560
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Summary                              "
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
            Left            =   90
            TabIndex        =   143
            Tag             =   "ShowSumm"
            ToolTipText     =   "Check To Show Summary"
            Top             =   2415
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Cust RmCd                            "
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
            Left            =   90
            TabIndex        =   142
            Tag             =   "ShowCustRm"
            ToolTipText     =   "Check To Show Customer Rm Code"
            Top             =   2130
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Cust DsgCd                            "
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
            Left            =   90
            TabIndex        =   141
            Tag             =   "ShowCustDsg"
            ToolTipText     =   "Check To Show Customer Design Code"
            Top             =   1845
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Order Remarks                    "
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
            Left            =   90
            TabIndex        =   139
            Tag             =   "ShowRem"
            ToolTipText     =   "Check To Show Order Master Level Remark"
            Top             =   1290
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Labour Value                                  "
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
            Left            =   90
            TabIndex        =   138
            Tag             =   "ShowLabVal"
            ToolTipText     =   "Check To Show Labour Value"
            Top             =   1005
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Value                                        "
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
            Index           =   14
            Left            =   90
            TabIndex        =   137
            Tag             =   "ShowRmVal"
            ToolTipText     =   "Check To Show Picture"
            Top             =   720
            Width           =   3600
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Design values                        "
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
            Left            =   90
            TabIndex        =   136
            Tag             =   "ShowDsgVal"
            ToolTipText     =   "Check To Show Design Value"
            Top             =   435
            Width           =   3600
         End
         Begin VB.Frame FraOptOrdQty 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   135
            ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
            Top             =   2700
            Width           =   1785
            Begin VB.OptionButton OptOrdQty 
               Caption         =   "Amd"
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
               TabIndex        =   51
               ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton OptOrdQty 
               Caption         =   " Enter"
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
               Left            =   900
               TabIndex        =   52
               ToolTipText     =   "Select Amended + Entered Qty O r Entered Qty"
               Top             =   0
               Width           =   885
            End
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Picture                                            "
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
            Left            =   90
            TabIndex        =   133
            Tag             =   "ShowPic"
            ToolTipText     =   "Check To Show Picture"
            Top             =   150
            Width           =   3600
         End
         Begin VB.Frame FraOptPtr 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   111
            ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
            Top             =   3600
            Width           =   1785
            Begin VB.OptionButton OptPtr 
               Caption         =   "Actual"
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
               Left            =   900
               TabIndex        =   61
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton OptPtr 
               Caption         =   "Calc"
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
               TabIndex        =   60
               ToolTipText     =   "Select Calculated Ptr Or Actual Pointer"
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.Frame FraOptDelDt 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2040
            TabIndex        =   110
            ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
            Top             =   3000
            Width           =   1785
            Begin VB.OptionButton OptDelDt 
               Caption         =   "    Prd"
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
               Left            =   900
               TabIndex        =   55
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
               Top             =   0
               Width           =   885
            End
            Begin VB.OptionButton OptDelDt 
               Caption         =   "Exp"
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
               ToolTipText     =   "Select Exp Del Date/ Prd Del Date"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FraOptRm 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2010
            TabIndex        =   109
            ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
            Top             =   3300
            Width           =   1755
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "    Ord"
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
               Left            =   900
               TabIndex        =   58
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   0
               Width           =   945
            End
            Begin VB.OptionButton OptRmPrd 
               Caption         =   "Prd"
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
               TabIndex        =   57
               ToolTipText     =   "Select Production Rm Qty, Wt or Order Rm Qty, Wt"
               Top             =   0
               Width           =   645
            End
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   5
            Left            =   3540
            TabIndex        =   43
            ToolTipText     =   "Show Price (Yes / No)"
            Top             =   435
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
            Index           =   14
            Left            =   3540
            TabIndex        =   45
            ToolTipText     =   "Show Labor Value (Yes / No)"
            Top             =   1005
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN4"
            IdName          =   "UOYN4"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   98
            Left            =   3540
            TabIndex        =   46
            ToolTipText     =   "Show Remarks in  Report? (Y/N)"
            Top             =   1290
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
            Index           =   35
            Left            =   3540
            TabIndex        =   47
            ToolTipText     =   "Show Cust DsgCd (Yes / No)"
            Top             =   1845
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
            Index           =   38
            Left            =   3540
            TabIndex        =   48
            ToolTipText     =   "Show Cust RmCd (Yes / No)"
            Top             =   2130
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
            Index           =   40
            Left            =   3540
            TabIndex        =   49
            ToolTipText     =   "Show Summary (Yes / No)"
            Top             =   2415
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN9"
            IdName          =   "UOYN9"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   52
            Left            =   3540
            TabIndex        =   44
            ToolTipText     =   "Show Rm Value (Yes / No)"
            Top             =   720
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN14"
            IdName          =   "UOYN14"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   6
            Left            =   3210
            TabIndex        =   56
            ToolTipText     =   "Show Production Rm Qty, Wt ('Y' For Prd/ 'N' For Ord)"
            Top             =   3300
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN3"
            ReCalcParent    =   "UOYN10"
            IdName          =   "UOYN3"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   19
            Left            =   3240
            TabIndex        =   53
            ToolTipText     =   "Show Export/ Production Delivery Date? (Enter ""Y"" for Export and ""N"" for Production)"
            Top             =   3000
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
            Index           =   41
            Left            =   3240
            TabIndex        =   59
            ToolTipText     =   "Show Calculated Pointer (Yes / No)"
            Top             =   3600
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN10"
            ReCalcOn        =   "UOYN3"
            IdName          =   "UOYN10"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   4
            Left            =   3540
            TabIndex        =   50
            ToolTipText     =   "Entered + Amendment Qty (Yes / No)"
            Top             =   2460
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
            Index           =   68
            Left            =   3540
            TabIndex        =   144
            ToolTipText     =   "Show Remarks in  Report? (Y/N)"
            Top             =   1560
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN26"
            IdName          =   "UOYN26"
         End
         Begin VB.Label lblAmdQty 
            BackStyle       =   0  'Transparent
            Caption         =   "Amended Qty"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   134
            Top             =   2700
            Width           =   1305
         End
         Begin VB.Label LblCalcPtr 
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
            Left            =   120
            TabIndex        =   108
            ToolTipText     =   "Location"
            Top             =   3600
            Width           =   1395
         End
         Begin VB.Label lblDelDt 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Del Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   107
            Top             =   3000
            Width           =   1440
         End
         Begin VB.Label lblPrdRmQtyWt 
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
            Left            =   120
            TabIndex        =   106
            Top             =   3300
            Width           =   1800
         End
         Begin VB.Label LblRmVal 
            BackStyle       =   0  'Transparent
            Caption         =   "Rm Value"
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
            Left            =   120
            TabIndex        =   105
            ToolTipText     =   "Location"
            Top             =   720
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label LblShowSumm 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Summary"
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
            Left            =   120
            TabIndex        =   102
            ToolTipText     =   "Location"
            Top             =   2415
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   101
            ToolTipText     =   "Location"
            Top             =   1815
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
            Left            =   120
            TabIndex        =   100
            ToolTipText     =   "Location"
            Top             =   2130
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblShowRem 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Remarks "
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
            Height          =   285
            Left            =   120
            TabIndex        =   97
            Top             =   1290
            Visible         =   0   'False
            Width           =   2310
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
            Left            =   120
            TabIndex        =   96
            ToolTipText     =   "Location"
            Top             =   150
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label LblShwVal 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Design Values"
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
            Left            =   120
            TabIndex        =   95
            ToolTipText     =   "Location"
            Top             =   435
            Visible         =   0   'False
            Width           =   2265
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
            Left            =   120
            TabIndex        =   94
            ToolTipText     =   "Location"
            Top             =   1005
            Visible         =   0   'False
            Width           =   1845
         End
      End
      Begin VB.Frame FraDispSc 
         Height          =   3345
         Left            =   4200
         TabIndex        =   130
         Top             =   6120
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Wt On RmCtg "
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
            Index           =   29
            Left            =   120
            TabIndex        =   146
            Tag             =   "DmHold"
            ToolTipText     =   "Check To Show RmCtg wise Details instead of RmStg"
            Top             =   3000
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Ord Qty                        "
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
            Index           =   27
            Left            =   120
            TabIndex        =   127
            Tag             =   "ShowOrdQty"
            ToolTipText     =   "Check To Show Ord Qty"
            Top             =   2445
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Col              "
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
            Index           =   15
            Left            =   120
            TabIndex        =   113
            Tag             =   "ShowDsgCol"
            ToolTipText     =   "Check To Show Color"
            Top             =   465
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Grs Wt         "
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
            Index           =   16
            Left            =   120
            TabIndex        =   115
            Tag             =   "ShowGrsWt"
            ToolTipText     =   "Check To Show Design Gross Wt"
            Top             =   750
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Code Details     "
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
            Index           =   17
            Left            =   120
            TabIndex        =   117
            Tag             =   "ShowRmCd"
            ToolTipText     =   "Check To Show Code Details"
            Top             =   1035
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm RmCtg/SCtg        "
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
            Index           =   18
            Left            =   120
            TabIndex        =   119
            Tag             =   "ShowRmSCtg"
            ToolTipText     =   "Check To Show Rm SubCtg Details"
            Top             =   1320
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Wt Details         "
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
            Index           =   19
            Left            =   120
            TabIndex        =   121
            Tag             =   "ShowRmWt"
            ToolTipText     =   "Check To Show Rm Weight Details"
            Top             =   1605
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Rm Quantity           "
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
            Index           =   20
            Left            =   120
            TabIndex        =   123
            Tag             =   "ShowRmQty"
            ToolTipText     =   "Check To Show Rm Quantity"
            Top             =   1890
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit Price               "
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
            Index           =   21
            Left            =   120
            TabIndex        =   125
            Tag             =   "ShowUnitPrc"
            ToolTipText     =   "Check To Show Unit Price"
            Top             =   2175
            Width           =   2445
         End
         Begin VB.CheckBox ChkBoxArr 
            Alignment       =   1  'Right Justify
            Caption         =   "Sub PO                   "
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
            Index           =   22
            Left            =   120
            TabIndex        =   129
            Tag             =   "ShowOdPO"
            ToolTipText     =   "Check To Show Sub PO"
            Top             =   2730
            Width           =   2445
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   55
            Left            =   2490
            TabIndex        =   120
            ToolTipText     =   "Show Rm Weight Details (Yes/No)?"
            Top             =   1590
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN19"
            ReCalcParent    =   "UOYN10"
            IdName          =   "UOYN19"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   56
            Left            =   2490
            TabIndex        =   118
            ToolTipText     =   "Show Rm Sub Category (Yes/No) ?"
            Top             =   1305
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN18"
            IdName          =   "UOYN18"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   57
            Left            =   2490
            TabIndex        =   122
            ToolTipText     =   "Show Rm Quantity (Yes/No) ?"
            Top             =   1875
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN20"
            IdName          =   "UOYN20"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   58
            Left            =   2490
            TabIndex        =   124
            ToolTipText     =   "Show Unit Price (Yes/No)?"
            Top             =   2160
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN21"
            IdName          =   "UOYN21"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   59
            Left            =   2490
            TabIndex        =   128
            ToolTipText     =   "Show Sub PO (Yes/No) ?"
            Top             =   2730
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN22"
            IdName          =   "UOYN22"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   61
            Left            =   2490
            TabIndex        =   116
            ToolTipText     =   "Show Rm Code Details (Yes / No)"
            Top             =   1020
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN17"
            IdName          =   "UOYN17"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   53
            Left            =   2490
            TabIndex        =   112
            ToolTipText     =   "Show Design Color? (Yes/No)"
            Top             =   465
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN15"
            IdName          =   "UOYN15"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   54
            Left            =   2490
            TabIndex        =   114
            ToolTipText     =   "Show Design Grs Wt? (Yes / No)"
            Top             =   750
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN16"
            IdName          =   "UOYN16"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   73
            Left            =   2490
            TabIndex        =   126
            ToolTipText     =   "Show Ord Qty (Yes/No) ?"
            Top             =   2445
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN27"
            IdName          =   "UOYN27"
         End
         Begin MwfCtl.MWCTL_MED atxt 
            Height          =   285
            Index           =   78
            Left            =   2490
            TabIndex        =   147
            ToolTipText     =   "Show RmCtg wise Wts (Y/N/' ')"
            Top             =   3000
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            Enabled         =   0   'False
            DataType        =   4
            DataField       =   "UOYN29"
            IdName          =   "UOYN29"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Show Design Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   0
            TabIndex        =   132
            Top             =   150
            Width           =   2910
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFgCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepOrd
Dim Rep1 As New EmrRepOrdCat
Dim moCn As MwfLib.MDOConnection    ' *** (Bef speed) bef 24/11/06
Dim ws_BaseCurCd As String, wb_InBaseCur As Boolean
Dim wChkSel As Integer
Dim wo_rsPic As MDORowSet
Private Sub SetGroupSort()
  'GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
On Error GoTo RepErr
  'SuperGems.2-OdDmId field removed, because giving error
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  Dim wExcelOrdBy As String
  
  If FraSC.Visible = True Then CmdSc_Click
  
  Dim wrepcnd As String, wCnd As String, wSqlStrg As String, ws_RmQty As String, ws_RmWt As String
  Dim wCommStr As String, wRmStr As String, wRmGrp As String, wRmCommStr As String
  Dim wLabStr As String, wLabGrp As String, wLabCommStr As String, wDsgCtgStr As String
  Dim wRmDet As String, wOdSalPrcOpt As String, wRmValOpt As String, wRmSetValOpt As String
  Dim wLabValOpt As String, wTcTyp As String, wDmTcTyp As String, wRmDetSql As String
  Dim wLabRtOpt As String, wDelDt As String, wDelStr As String, ws_RmPtr As String
  Dim wRtStr As String 'urmi 2.04-3 to get the rate when grouping is by RMCd (TotVal/Wt)
  Dim wInGrp2 As String       ' **** Manali 3.6.0
  
  Dim wOurDsgCd As String, wOurRmCd As String
  Dim wOrdCTCommStr As String
  Dim ms_TcDesc As String
  
  Set Rep = Nothing
  Set Rep1 = Nothing

    ADC("UoCoCdFr") = gs_CoCd
    wDmTcTyp = "DM"
  
    GRP_REP.SetFormula Rep1, "wGrpOpt", ""

    Set ADC.RepSource = Rep1
    Set mRep = Rep1
  
    mRep.TxtHead.SetText "FG Catalogue"
  
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  If ADC("UoCurCdFr") = "" Then
    wb_InBaseCur = True
    mRep.TxtValIn.SetText "Value In " + ws_BaseCurCd
    'If UCase(adc.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + ws_BaseCurCd + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Price" + " (" + ws_BaseCurCd + ")"
  Else
    wb_InBaseCur = False
    mRep.TxtValIn.SetText "Value In " + ADC("UoCurCdFr")
    'If UCase(adc.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Unit Price" + vbCrLf + "(" + adc("UoCurCdFr") + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur2.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
    If UCase(ADC.MenuCd) = UCase("RepOrdCat") Then mRep.TxtUnitPrcCur3.SetText "Price" + " (" + ADC("UoCurCdFr") + ")"
  End If
  
    Call DispCoNm
    Call SetMwName(mRep)
      
    wrepcnd = ADC.RepCond
    wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
    wCnd = wCnd + IIF(wb_InBaseCur = True, "", " and OmCmCurCd= '" + ADC("UoCurCdFr") + "' ")
   
     wOdSalPrcOpt = IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc")
    
      wRmValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")"
    
    If (ADC("UoYN4")) = "Y" Then
        wRmSetValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OrSetSalVal/OmCnvFct else OrSetSalVal*OmCnvFct end) ", "OrSetSalVal") + ") as OrSetSalVal"
        wLabValOpt = "sum(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalVal/OmCnvFct else OlSalVal*OmCnvFct end) ", "OlSalVal") + ") as OlSalVal"
        wLabRtOpt = "max(" + IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OlSalRt/OmCnvFct else OlSalRt*OmCnvFct end) ", "OlSalRt") + ") as OlSalRt"
        If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'N'"
    Else
      wRmSetValOpt = "0.00 as OrSetSalVal": wLabValOpt = "0.00 as OlSalVal": wLabRtOpt = "0.00 as OlSalRt"
      If moCn.RecSeek("Select PmCd from Param where PTyp = 'SYS' and PmCd = 'ECH'") <> False Then GRP_REP.SetFormula Rep, "wLabRtYN", "'Y'"
    End If
          
    
    If ADC("UoYN3") = "Y" Then
      ws_RmQty = "OrPrdQty": ws_RmWt = "OrPrdWt"
    Else
      ws_RmQty = "OrQty": ws_RmWt = "OrWt"
    End If
    
   ' wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
   ' wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
    
      If UCase(ADC("UoYN7")) = "Y" Then
        wOurDsgCd = "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                          "Where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                          "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                          "CdOurDmSz = OdDmSz), " + _
                        "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                              "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                              "CdOurDmCd = OdDmCd and CdOurSfx = OdSfx and " + _
                              "CdOurDmSz = ''), " + _
                            "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                  "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                                  "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                                  "CdOurDmSz = OdDmSz), " + _
                                "IsNull((Select cast(CdCustDmCd as Varchar(17)) from CustDsg " + _
                                      "where " + IIF(gs_Partition = ctCurrPrtn, " CdPrtKey='" + ctCurrPrtn + "' and ", "") + " CdCtg = 'C' and CdCd = OmCmCd and " + _
                                      "CdOurDmCd = OdDmCd and CdOurSfx = '' and " + _
                                      "CdOurDmSz = ''), '{' + OdDmCd + '}')))) as CdOurDmCd"
      Else
        wOurDsgCd = "OdDmCd as CdOurDmCd"
      End If
    
      wOurRmCd = "Max(OrRmCd) as OrRmCd"
    
    Dim ws_BaseRefJoin As String, ws_BaseRefFlds As String, ws_BaseRef As String
    ws_BaseRefJoin = "": ws_BaseRef = ""
    
    Dim wGrpArr1() As String, wGrpFormula1 As String, wGrpDescFormula1 As String
    Dim wGrpArr2() As String, wGrpFormula2 As String, wGrpDescFormula2 As String
    Dim i As Integer, qDmCtgDesc As String, qRmCtgDesc As String, qDmColDesc As String
    Dim qVaCtgDesc As String, qMainMet As String, wOrJoin As String
    Dim qDmSalCtgDesc As String
    Dim wBagClsCnd  As String
    
    qDmCtgDesc = "'' as qDmCtgDesc": qRmCtgDesc = "'' as qRmCtgDesc"
    qDmColDesc = "'' as qDmColDesc": qVaCtgDesc = "'' as qVaCtgDesc"
    qDmSalCtgDesc = "'' as qDmSalCtgDesc"
    
    qMainMet = "space(2) as OrRmCtg": wOrJoin = ""
  
    wGrpArr1 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + ADC("UoInGrp") + "' "), ",")
    For i = 0 To UBound(wGrpArr1)
      Select Case UCase(wGrpArr1(i))
       Case Is = "RMCTG"
            qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                       "(select RmCtg from RmMst,OrdRm where RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + "  and  OrMainMet='Y'  " + _
                       "and OrCocd = OmCocd And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr  " + _
                       "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OdPrtKey ", "") + ")),'') as qRmCtgDesc"
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
       Case Is = "KT"
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.IdKt}"
       Case Is = "DMCOL"
            qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
       Case Is = "VACTG"
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= OdVaCtg) as qVaCtgDesc"
            wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
       Case Is = "PCS/PRS"
          wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "{rdo.DmUom}"
       Case Is = "DMCTG"
            qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd=DmCtg) as qDmCtgDesc"
          wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
      Case Is = "GLDLS"
           wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
      Case Is = "SALCTG"
          qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
        wGrpDescFormula1 = wGrpDescFormula1 + IIF(wGrpDescFormula1 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
      End Select
      wGrpFormula1 = wGrpFormula1 + IIF(wGrpFormula1 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'") + "})))"
      
           wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + moCn.GetFldVal("Select (Case PDesc225 when 'IdKt' Then 'OdKT' when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=OdVaCtg and PSCd='''')' " + _
                       " When 'IdVaCtg' Then 'OdVaCtg' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                       " OrMainMet=''Y'' and OrCocd = OmCocd And OrTc = OmTc and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'''')' Else PDesc225 End)   " + _
                       " from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr1(i) + "'")
    Next i
  
    wGrpArr2 = Split(moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGR' and " + _
           "PMCd= '" + ADC("UoInGrp1") + "' "), ",")
    For i = 0 To UBound(wGrpArr2)
     Select Case UCase(wGrpArr2(i))
      Case Is = "RMCTG"
           qRmCtgDesc = "IsNull((Select PDesc from Param where PTyp= 'RMCTG' and PMCd= " + _
                     "(select RmCtg from RmMst,OrdRm where RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + "  and  OrMainMet='Y' " + _
                     "and OrCocd = OmCocd And OrTc = OmTc  and OrYy = OmYy and OrChr=OmChr " + _
                     "and OrNo=OmNo and OrSr=OdSr " + IIF(gs_Partition = ctCurrPrtn, " and OrPrtKey=OmPrtKey ", "") + ")),'') as qRmCtgDesc"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qRmCtgDesc})"
     Case Is = "KT"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.IdKt}"
     Case Is = "DMCOL"
          qDmColDesc = "(Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qDmColDesc"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmColDesc})"
     Case Is = "VACTG"
            qVaCtgDesc = "(Select PDesc from Param where PTyp= 'VACTG' and PMCd= OdVaCtg) as qVaCtgDesc"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qVaCtgDesc})"
    Case Is = "PCS/PRS"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "{rdo.DmUom}"
    Case Is = "DMCTG"
          qDmCtgDesc = " (Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) as qDmCtgDesc"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmCtgDesc})"
    Case Is = "GLDLS"
        wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "'@'+ToText({rdo.IdGldLs}, '##0.000')+ '%'"
    Case Is = "SALCTG"
        qDmSalCtgDesc = " (Select PDesc from Param where PTyp= 'SALCTG' and PMCd= DmCtg and PSCd= DmSalCtg) as qDmSalCtgDesc "
      wGrpDescFormula2 = wGrpDescFormula2 + IIF(wGrpDescFormula2 = "", "", "+', '+") + "Trim({rdo.qDmSalCtgDesc})"
    End Select
    wGrpFormula2 = wGrpFormula2 + IIF(wGrpFormula2 = "", "", "+") + "Trim(UpperCase(ToText({rdo." + moCn.GetFldVal("Select PDesc225 from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'") + "})))"
     
        wInGrp2 = moCn.GetFldVal("Select (Case PDesc225 when 'IdKt' Then 'OdKT' when 'IdGldLS' " + _
                       " Then '(Select PNum from Param where PTyp=''VACTG'' and PMCd=OdVaCtg and PSCd='''')' " + _
                       " When 'IdVaCtg' Then 'OdVaCtg' " + _
                       " When 'OrRmCtg' Then 'IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                       " OrMainMet=''Y'' and OrCocd = OmCocd And OrTc = OmTc and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'''')' Else PDesc225 End)  from Param where PTyp= 'INGRPRCP' and PMCd= '" + wGrpArr2(i) + "'")
     wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy <> "" And InStr(1, wExcelOrdBy, wInGrp2) = 0, ",", "") + IIF(InStr(1, wExcelOrdBy, wInGrp2) = 0, wInGrp2, "")
     
    Next i
    
    
    
    If ADC("UOYN11") = "Y" Then
      GRP_REP.SetFormula mRep, "OdSr", "Trim(UpperCase(ToText({rdo.OdDmCd})))+ '/'+ ToText ({rdo.OdSr}, '####0') "
      
      wExcelOrdBy = wExcelOrdBy + IIF(wExcelOrdBy = "", "", ",") + " OdDmCd "
    Else
      GRP_REP.SetFormula mRep, "OdSr", "ToText ({rdo.OdSr}, '####0') "
    End If
    
    wBagClsCnd = IIF(ADC("UoYN") = "Y", " and BCls= 'N' ", IIF(ADC("UoYN") = "N", " and BCls= 'Y' ", ""))
    
                 
      wSqlStrg = "Select 0 as qRow, 0 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr ,Space(17) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx, DmCtg, DmSalCtg," + _
               "DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ,PNum as IdGldLs, " + _
               "Space(254) as qRmCtgDesc,Space(254) as  qDmCtgDesc,Space(254) as  qDmSalCtgDesc,Space(254) as  qDmColDesc,Space(254) as  qVaCtgDesc, " + _
               "OrRmCtg  as OrRmctg ,OdGldAsWt, OdLabAsWt,  OdSalPrc  as OdSalPrc, OdOrdEnt as qOrdQty, Space(254) as qColDesc, " + _
               "OrPrdWt as qGrsWt, 0 as OdTrayNo   Into #TmpA " + _
               " From OrdMst,OrdDsg,Param,CustMst,DsgMst,OrdRm  where 1=2 "
      moCn.CreateTmpTable "#TmpA", wSqlStrg
      
      wSqlStrg = "Select 0 as qRow, 0 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr ,Space(17) as CdOurDmCd, OdDmCd, OdDmSz, OdSfx, DmCtg, DmSalCtg," + _
               "DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ,PNum as IdGldLs, " + _
               "Space(254) as qRmCtgDesc,Space(254) as  qDmCtgDesc,Space(254) as  qDmSalCtgDesc,Space(254) as  qDmColDesc,Space(254) as  qVaCtgDesc, " + _
               "OrRmCtg  as OrRmctg ,OdGldAsWt, OdLabAsWt,  OdSalPrc  as OdSalPrc, OdOrdEnt as qOrdQty, Space(254) as qColDesc, " + _
               "OrPrdWt as qGrsWt, 0 as OdTrayNo   Into #TmpB " + _
               " From OrdMst,OrdDsg,Param,CustMst,DsgMst,OrdRm  where 1=2 "
      moCn.CreateTmpTable "#TmpB", wSqlStrg
      
      
      wExcelOrdBy = IIF(wExcelOrdBy <> "", ",", "") + wExcelOrdBy
      
      
      Dim wFmCnd As String
      wFmCnd = IIF(ADC("UoTTcFr") <> "", " FmdTc>='" + ADC("UoTTcFr") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTTcTo") <> "" And wFmCnd <> "", " And ", "") + IIF(ADC("UoTTcTo") <> "", " FmdTc<='" + ADC("UoTTcTo") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTYyFr") <> "" And wFmCnd <> "", " And ", "") + IIF(ADC("UoTYyFr") <> "", " FmdYy>='" + ADC("UoTYyFr") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTYyTo") <> "" And wFmCnd <> "", " And ", "") + IIF(ADC("UoTYyTo") <> "", " FmdYy<='" + ADC("UoTYyTo") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTChrFr") <> "" And wFmCnd <> "", " And ", "") + IIF(ADC("UoTChrFr") <> "", " FmdChr>='" + ADC("UoTChrFr") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTChrTo") <> "" And wFmCnd <> "", " And ", "") + IIF(ADC("UoTChrTo") <> "", " FmdChr<='" + ADC("UoTChrTo") + "'", "")
      wFmCnd = wFmCnd + IIF(ADC("UoTNoFr") <> 0 And wFmCnd <> "", " And ", "") + IIF(ADC("UoTNoFr") <> 0, " FmdNo>=" + CStr(ADC("UoTNoFr")), "")
      wFmCnd = wFmCnd + IIF(ADC("UoTNoTo") <> 0 And wFmCnd <> "", " And ", "") + IIF(ADC("UoTNoTo") <> 0, " FmdNo<=" + CStr(ADC("UoTNoTo")), "")
      wFmCnd = wFmCnd + IIF(ADC("UoTdSrFr") <> 0 And wFmCnd <> "", " And ", "") + IIF(ADC("UoTdSrFr") <> 0, " FmdSr>=" + CStr(ADC("UoTdSrFr")), "")
      wFmCnd = wFmCnd + IIF(ADC("UoTdSrTo") <> 0 And wFmCnd <> "", " And ", "") + IIF(ADC("UoTdSrTo") <> 0, " FmdSr<=" + CStr(ADC("UoTdSrTo")), "")
      
      If wFmCnd <> "" Then
            wCnd = wCnd + IIF(wCnd = "", " ", " And ") + "Exists(select 'x' From Fmd Where FmdCoCd=FdCoCd and FdYy=FmdFdYy and FdKey=FmdFdKey and FmdCoCd='" + ADC("UOCOCDFR") + "' and " + wFmCnd + ")"
      End If
        
        
    Dim wDstJoin As String
    If ADC("UOYN28") = "Y" Then
     wDstJoin = "Join (Select Max(OmCoCd+OmTc+OmYy+OmChr+Str(OmNo)+Str(OdSr)) as qOdSr,OdDmCd as qDmCd  From #TmpB  Fd " + _
                " where  OmDt=(select max(OmDt)  From #TmpB Fd1 Where Fd.OdDmCd=Fd1.OdDmCd  Group By OdDmCd) Group By OdDmCd) Fd " + _
                "  On FdCoCd+FdTc+FdYy+FdChr+Str(FdNo)+Str(FdSr)=qOdSr  "
    End If

  
      wSqlStrg = "Insert Into #TmpB " + _
                 "Select 0 as qRow,0 as qCol,FgCoCd, FgYy, FgTc, FgChr, FgNo, OmCmCd,FgDt as OmDt, OmPoNo, FdSr ," + wOurDsgCd + ", OdDmCd, OdDmSz, OdSfx," + _
                 "DmCtg, DmSalCtg, DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ," + _
                 "PNum as IdGldLs" + _
                 "," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qVaCtgDesc + "," + _
                 "IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                 "OrMainMet='Y' and OrCocd = OmCocd And OrTc = OmTc  " + _
                 "and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'') as OrRmctg /*qMainMetRmCtg*/," + _
                 "OdGldAsWt, OdLabAsWt, " + _
                 IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " as OdSalPrc, " + _
                 "FdQty as qOrdQty, (Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
                 "FdGrWt as qGrsWt, 0 as OdTrayNo  " + _
                 " From Fg Join Fgd On FgCoCd=FdCoCd and FgTc=FdTc and FgYy=FgYy and FgChr=FdChr and FgNo=FdNo " + _
                 " Join Bag On FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo  " + _
                 " join OrdDsg on OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy  and OdChr= FdExpOdChr and OdNo= FdExpOdNo and OdSr=FdExpOdSr and OdPrtKey=FdPrtKey" + _
                 " Join OrdMst on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy  and OdChr= OmChr and OdNo= OmNo and OdPrtKey=OdPrtKey" + _
                 " Join Param a On a.PTyp='VACTG' and a.PMCd=OdVaCtg and a.PSCd=''" + _
                 " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                 " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                 wCnd + wBagClsCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "")
                 
                 moCn.Execute (wSqlStrg)
                 
                 
            '     wSqlStrg = "Insert Into #TmpA " + _
                 "Select (row_Number() over(Partition By FgIdNo Order By FgCoCd,FgTc,FgYy,FgChr,FgNo" + wExcelOrdBy + ",FdSr)-1)/3 as qRow," + _
                 " row_Number() over(Partition By FgIdNo Order By FgCoCd,FgTc,FgYy,FgChr,FgNo" + wExcelOrdBy + ",FdSr)%3 as qCol,OmCoCd, OmYy, OmTc, OmChr, OmNo, OmCmCd, OmDt, OmPoNo, OdSr ,CdOurDmCd, OdDmCd, OdDmSz, OdSfx, OdDmId, DmCtg, DmSalCtg," + _
                "DmTcTyp, OdDmCol, OdPicNm, OdPoNo,IdKt,IdVaCtg, DmUom ,IdGldLs, " + _
                "qRmCtgDesc,qDmCtgDesc,qDmSalCtgDesc,qDmColDesc,qVaCtgDesc, " + _
                "OrRmctg ,OdGldAsWt, OdLabAsWt,OdSalPrc,qOrdQty,qColDesc,OrPrdWt as qGrsWt, 0 as OdTrayNo " + _
                " From #TmpB " + wDstJoin
                
                
                wSqlStrg = "Insert Into #TmpA " + _
                       "Select (row_Number() over(" + IIF(ADC("UOYN28") = "Y", " ", "Partition By FgIdNo ") + " Order By FgCoCd,FgTc,FgYy,FgChr,FgNo" + wExcelOrdBy + ",FdSr)-1)/3 as qRow," + _
                        " row_Number() over(" + IIF(ADC("UOYN28") = "Y", " ", "Partition By FgIdNo ") + " Order By FgCoCd,FgTc,FgYy,FgChr,FgNo" + wExcelOrdBy + ",FdSr)%3 as qCol, " + _
                        "FgCoCd, FgYy, FgTc, FgChr, FgNo, OmCmCd,FgDt as OmDt, OmPoNo," + IIF(ADC("UOYN28") = "Y", " FdIdNo as ", "FdSr as ") + " FdSr," + wOurDsgCd + ", OdDmCd, OdDmSz, OdSfx," + _
                        "DmCtg, DmSalCtg, DmTcTyp, OdDmCol, OdPicNm, OdPoNo, OdKt as IdKt, OdVaCtg as IdVaCtg, DmUom ," + _
                        "PNum as IdGldLs" + _
                        "," + qRmCtgDesc + "," + qDmCtgDesc + "," + qDmSalCtgDesc + "," + qDmColDesc + "," + qVaCtgDesc + "," + _
                        "IsNull((select max(RmCtg) from RmMst ,OrdRm where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey=OrPrtKey and ", "") + " RmCd=OrRmCd  and  " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey=OmPrtKey and ", "") + _
                        "OrMainMet='Y' and OrCocd = OmCocd And OrTc = OmTc  " + _
                        "and OrYy = OmYy and OrChr=OmChr and OrNo=OmNo and OrSr=OdSr),'') as OrRmctg /*qMainMetRmCtg*/," + _
                        "OdGldAsWt, OdLabAsWt, " + _
                        IIF(wb_InBaseCur, " (Case when OmMulDiv= 'M' then OdSalPrc/OmCnvFct else OdSalPrc*OmCnvFct end) ", "OdSalPrc") + " as OdSalPrc, " + _
                        "FdQty as qOrdQty, (Select PDesc from Param where PTyp= 'DMCOL' and PMCd= OdDmCol) as qColDesc, " + _
                        "FdGrWt as qGrsWt, 0 as OdTrayNo  " + _
                        " From Fg Join Fgd On FgCoCd=FdCoCd and FgTc=FdTc and FgYy=FgYy and FgChr=FdChr and FgNo=FdNo " + _
                        " Join Bag On FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo  " + _
                        " join OrdDsg on OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy  and OdChr= FdExpOdChr and OdNo= FdExpOdNo and OdSr=FdExpOdSr and OdPrtKey=FdPrtKey" + _
                        " Join OrdMst on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy  and OdChr= OmChr and OdNo= OmNo and OdPrtKey=OdPrtKey" + _
                        " Join Param a On a.PTyp='VACTG' and a.PMCd=OdVaCtg and a.PSCd=''" + _
                        " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
                        " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= ''  " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
                         wDstJoin + wCnd + wBagClsCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "")
                 
                
                 
                 '(index= OrdDsg)
    moCn.Execute (wSqlStrg)
    
  
    
    wSqlStrg = " Select  OmCoCd," + IIF(ADC("UOYN28") = "Y", "'' as OmYy, OmTc,'' as OmChr,0 as OmNo ", "OmYy, OmTc, OmChr, OmNo") + " ,qRow,Max(OmCmCd) as OmCmCd ,Max(OmDt) as OmDt,Max(OmPoNo) as OmPoNo, " + _
               " Max((Case when qCol=1 Then   OdSr else 0 End)) as OdSr1,Max((Case when qCol=1 Then   CdOurDmCd else '' End)) as CdOurDmCd1, " + _
               " Max((Case when qCol=1 Then   OdDmCd else '' End)) as OdDmCd1,Max((Case when qCol=1 Then   OdDmSz else '' End)) as OdDmSz1, " + _
               " Max((Case when qCol=1 Then   OdSfx else '' End)) as OdSfx1, " + _
               " Max((Case when qCol=1 Then   DmCtg else '' End)) as DmCtg1,Max((Case when qCol=1 Then   DmSalCtg else '' End)) as DmSalCtg1, " + _
               " Max((Case when qCol=1 Then   DmTcTyp else '' End)) as DmTcTyp1,Max((Case when qCol=1 Then   OdDmCol else '' End)) as OdDmCol1, " + _
               " Max((Case when qCol=1 Then   OdPicNm else '' End)) as OdPicNm1,Max((Case when qCol=1 Then   OdPoNo else '' End)) as OdPoNo1, " + _
               " Max((Case when qCol=1 Then   IdKt else '' End)) as IdKt1,Max((Case when qCol=1 Then   IdVaCtg else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   DmUom else '' End)) as DmUom1,Max((Case when qCol=1 Then   IdGldLs else '' End)) as IdVaCtg1, " + _
               " Max((Case when qCol=1 Then   qRmCtgDesc else '' End)) as qRmCtgDesc1,Max((Case when qCol=1 Then   qDmCtgDesc else '' End)) as qDmCtgDesc1, " + _
               " Max((Case when qCol=1 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc1,Max((Case when qCol=1 Then   qDmColDesc else '' End)) as qDmColDesc1," + _
               " Max((Case when qCol=1 Then   qVaCtgDesc else '' End)) as qVaCtgDesc1,Max((Case when qCol=1 Then   OrRmctg else '' End)) as OrRmctg1, " + _
               " Max((Case when qCol=1 Then   OdGldAsWt else 0 End)) as OdGldAsWt1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   OdSalPrc else 0 End)) as OdSalPrc1,Max((Case when qCol=1 Then   OdLabAsWt else 0 End)) as OdLabAsWt1, " + _
               " Max((Case when qCol=1 Then   qOrdQty else 0 End)) as qOrdQty1,Max((Case when qCol=1 Then   qColDesc else '' End)) as qColDesc1, " + _
               " Max((Case when qCol=1 Then   qGrsWt else 0 End)) as qGrsWt1,Max((Case when qCol=1 Then   OdTrayNo else '' End)) as OdTrayNo1, "
      wSqlStrg = wSqlStrg + "Max((Case when qCol=2 Then   OdSr else 0 End)) as OdSr2,Max((Case when qCol=2 Then   CdOurDmCd else '' End)) as CdOurDmCd2, " + _
                        " Max((Case when qCol=2 Then   OdDmCd else '' End)) as OdDmCd2,Max((Case when qCol=2 Then   OdDmSz else '' End)) as OdDmSz2," + _
                        " Max((Case when qCol=2 Then   OdSfx else '' End)) as OdSfx2," + _
                        " Max((Case when qCol=2 Then   DmCtg else '' End)) as DmCtg2,Max((Case when qCol=2 Then   DmSalCtg else '' End)) as DmSalCtg2, " + _
                        " Max((Case when qCol=2 Then   DmTcTyp else '' End)) as DmTcTyp2,Max((Case when qCol=2 Then   OdDmCol else '' End)) as OdDmCol2, " + _
                        " Max((Case when qCol=2 Then   OdPicNm else '' End)) as OdPicNm2,Max((Case when qCol=2 Then   OdPoNo else '' End)) as OdPoNo2, " + _
                        " Max((Case when qCol=2 Then   IdKt else '' End)) as IdKt2,Max((Case when qCol=2 Then   IdVaCtg else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   DmUom else '' End)) as DmUom2,Max((Case when qCol=2 Then   IdGldLs else '' End)) as IdVaCtg2, " + _
                        " Max((Case when qCol=2 Then   qRmCtgDesc else '' End)) as qRmCtgDesc2,Max((Case when qCol=2 Then   qDmCtgDesc else '' End)) as qDmCtgDesc2, " + _
                        " Max((Case when qCol=2 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc2,Max((Case when qCol=2 Then   qDmColDesc else '' End)) as qDmColDesc2, " + _
                        " Max((Case when qCol=2 Then   qVaCtgDesc else '' End)) as qVaCtgDesc2,Max((Case when qCol=2 Then   OrRmctg else '' End)) as OrRmctg2, " + _
                        " Max((Case when qCol=2 Then   OdGldAsWt else 0 End)) as OdGldAsWt2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2, " + _
                        " Max((Case when qCol=2 Then   OdSalPrc else 0 End)) as OdSalPrc2,Max((Case when qCol=2 Then   OdLabAsWt else 0 End)) as OdLabAsWt2," + _
                        " Max((Case when qCol=2 Then   qOrdQty else 0 End)) as qOrdQty2,Max((Case when qCol=2 Then   qColDesc else '' End)) as qColDesc2, " + _
                        " Max((Case when qCol=2 Then   qGrsWt else 0 End)) as qGrsWt2,Max((Case when qCol=2 Then   OdTrayNo else '' End)) as OdTrayNo2, "
      wSqlStrg = wSqlStrg + " Max((Case when qCol=0 Then   OdSr else 0 End)) as OdSr3,Max((Case when qCol=0 Then   CdOurDmCd else '' End)) as CdOurDmCd3, " + _
                            " Max((Case when qCol=0 Then   OdDmCd else '' End)) as OdDmCd3,Max((Case when qCol=0 Then   OdDmSz else '' End)) as OdDmSz3, " + _
                            " Max((Case when qCol=0 Then   OdSfx else '' End)) as OdSfx3," + _
                            " Max((Case when qCol=0 Then   DmCtg else '' End)) as DmCtg3,Max((Case when qCol=0 Then   DmSalCtg else '' End)) as DmSalCtg3, " + _
                            " Max((Case when qCol=0 Then   DmTcTyp else '' End)) as DmTcTyp3,Max((Case when qCol=0 Then   OdDmCol else '' End)) as OdDmCol3, " + _
                            " Max((Case when qCol=0 Then   OdPicNm else '' End)) as OdPicNm3,Max((Case when qCol=0 Then   OdPoNo else '' End)) as OdPoNo3, " + _
                            " Max((Case when qCol=0 Then   IdKt else '' End)) as IdKt3,Max((Case when qCol=0 Then   IdVaCtg else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   DmUom else '' End)) as DmUom3,Max((Case when qCol=0 Then   IdGldLs else '' End)) as IdVaCtg3, " + _
                            " Max((Case when qCol=0 Then   qRmCtgDesc else '' End)) as qRmCtgDesc3,Max((Case when qCol=0 Then   qDmCtgDesc else '' End)) as qDmCtgDesc3, " + _
                            " Max((Case when qCol=0 Then   qDmSalCtgDesc else '' End)) as qDmSalCtgDesc3,Max((Case when qCol=0 Then   qDmColDesc else '' End)) as qDmColDesc3, " + _
                            " Max((Case when qCol=0 Then   qVaCtgDesc else '' End)) as qVaCtgDesc3,Max((Case when qCol=0 Then   OrRmctg else '' End)) as OrRmctg3, " + _
                            " Max((Case when qCol=0 Then   OdGldAsWt else 0 End)) as OdGldAsWt3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   OdSalPrc else 0 End)) as OdSalPrc3,Max((Case when qCol=0 Then   OdLabAsWt else 0 End)) as OdLabAsWt3, " + _
                            " Max((Case when qCol=0 Then   qOrdQty else 0 End)) as qOrdQty3,Max((Case when qCol=0 Then   qColDesc else '' End)) as qColDesc3, " + _
                            " Max((Case when qCol=0 Then   qGrsWt else 0 End)) as qGrsWt3,Max((Case when qCol=0 Then   OdTrayNo else '' End)) as OdTrayNo3 " + _
                            "  From #TmpA " + _
                            " Group By  OmCoCd,OmTc" + IIF(ADC("UOYN28") = "Y", "", ",OmYy,OmChr,OmNo") + ",qRow Order By OmCoCd,OmYy,OmTc,OmChr,OmNo,qRow "
     
      '****** Sachin 3.02 - Id fields In Joins
      ADC.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
      
      ' ***** Manali 3.03 -17/06/08 - Pd Mod - 'L' added
      ' ***** Manali Trading Module - CmCtg='C' replaced by gs_CmCtg
      wRmDet = "Select  " + IIF(ADC("UOYN28") = "Y", "'' as OdYy,FdTc as OdTc,'' as OdChr,0 as OdNo, FdIdNo as OdSr", "FdYy as OdYy,FdTc as OdTc,FdChr as OdChr,FdNo as OdNo, FdSr as OdSr") + ", 'N' as  OrMainMet, " + _
               "FrRmCtg as OrRmCtg, (case FrRmCtg when 'G' then 1 when 'P' then 2 when 'S' then 8 when 'L' then 9 when 'D' then 3 " + _
               "when 'C' then 4 when 'X' then 5 when 'A' then 6 when 'M' then 7 else 0 end) " + _
               "as qTag, FrRmSCtg as OrRmSCtg, FrRmCd as OrRmCd, FrRmQty as OrQty, FrRmWt as OrWt " + _
               "From Fgd " + _
               " join Fg on FdFgIdNo=FgIdNo and FdCoCd= FgCoCd and FdTc= FgTc and FdYy= FgYy and FdChr= FgChr and FdNo= FgNo " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=FgPrtKey ", "") + _
               " join FgRm on FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and " + _
               " FrNo= FdNo and FrSr= FdSr " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=FdPrtKey ", "") + _
               " Join Bag On FdCoCd=BCoCd and FdBYy=BYy and FdBChr=BChr and FdBNo=BNo  " + _
               " join OrdDsg on OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy  and OdChr= FdExpOdChr and OdNo= FdExpOdNo and OdSr=FdExpOdSr and OdPrtKey=FdPrtKey" + _
               " Join OrdMst on OdCoCd= OmCoCd and OdTc= OmTc and OdYy= OmYy  and OdChr= OmChr and OdNo= OmNo and OdPrtKey=OdPrtKey" + _
               " join CustMst on CmCtg='" + gs_CmCtg + "' and CmCd= OmCmCd " + _
               " join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= '" + wDmTcTyp + "' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=DmPrtKey ", "") + _
               wCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "")
               '(index= OrdRm)
        
      Call moCn.RepRes(mRep.Subreport1.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      Call moCn.RepRes(mRep.Subreport2.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      Call moCn.RepRes(mRep.Subreport3.OpenSubreport, wRmDet + ctMaxDopOpt)  '****** Sachin 3.01 added ctMaxDopOpt
      
      Set wo_rsPic = moCn.OpenRes("Select DmTcTyp as DmTcTyp,OdDmCd as DmCd,OdDmSz as DmSz,OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr,OdPicNm as PicNm from #TmpA" + _
                                  " order by OmCoCd, OmYy, OmTc, OmChr, OmNo,qRow,(Case when qCol=0 Then 3 Else qCol End)")
      'If wo_rsPic.RecCount > 0 Then CmdChgImg.Visible = True
   
    
    GRP_REP.SetFormula mRep, "wOrLmeSal", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                      " and Round({rdo.OrLmeSal}, 3) <> 0.00 and '" + ADC("UoYn14") + "'='Y' Then Round({rdo.OrLmeSal}, 3) else 0.00 "
    
    GRP_REP.SetFormula mRep, "wLnSv", " If ({rdo.RmCtg} ='G' or {rdo.RmCtg} ='P' or {rdo.RmCtg} ='S' or {rdo.RmCtg} ='L') " + _
                                      " then (if Round({rdo.OrLmeSal}, 3) = 0.00 Then Trim (ToText ({rdo.OrLn1}, 3)) )  " + _
                                      " else (If Round({rdo.OrLn1}, 3)= 0.01 Then '+0 ' " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.02 Then '+00 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0.03 Then '+000 '  " + _
                                      " else if Round({rdo.OrLn1}, 3)= 0 Then ' '  " + _
                                      " else Trim (ToText ({rdo.OrLn1}, 3)) ) "
    
    If ADC("UOYN28") = "Y" Then
        GRP_REP.SetFormula mRep, "wOrder", ""
        mRep.PhOrd.Suppress = True
        mRep.PHGRP.Suppress = True
   End If
      
      If ADC("UoYN15") = "Y" Or ADC("UoYN24") = "Y" Then
          Rep1.DetOdCol.Suppress = False
          If ADC("UoYN15") = "N" And ADC("UoYN24") = "Y" Then
            Rep1.lblOdCol.Suppress = True: Rep1.lblColOdCol.Suppress = True: Rep1.fldOdDmCol.Suppress = True
            Rep1.FldOdSr.Suppress = False: Rep1.lblOdSr.Suppress = False: Rep1.lblColOdSr.Suppress = False
            Rep1.FldOdSr2.Suppress = False: Rep1.lblOdSr2.Suppress = False: Rep1.lblColOdSr2.Suppress = False
            Rep1.FldOdSr3.Suppress = False: Rep1.lblOdSr3.Suppress = False: Rep1.lblColOdSr3.Suppress = False
            'Rep1.lblOdSr.Left = 70: Rep1.lblColOdSr.Left = 1090: Rep1.FldOdSr.Left = 1240
            
          ElseIf Not (ADC("UoYN15") = "Y" And ADC("UoYN24") = "Y") Then
            Rep1.FldOdSr.Suppress = True: Rep1.lblOdSr.Suppress = True: Rep1.lblColOdSr.Suppress = True
            Rep1.FldOdSr2.Suppress = True: Rep1.lblOdSr2.Suppress = True: Rep1.lblColOdSr2.Suppress = True
            Rep1.FldOdSr3.Suppress = True: Rep1.lblOdSr3.Suppress = True: Rep1.lblColOdSr3.Suppress = True
            
            Rep1.lblOdCol.Suppress = False: Rep1.lblColOdCol.Suppress = False: Rep1.fldOdDmCol.Suppress = False
            Rep1.lblOdCol2.Suppress = False: Rep1.lblColOdCol2.Suppress = False: Rep1.fldOdDmCol2.Suppress = False
            Rep1.lblOdCol3.Suppress = False: Rep1.lblColOdCol3.Suppress = False: Rep1.fldOdDmCol3.Suppress = False
          End If
      Else
        Rep1.DetOdCol.Suppress = True
      End If

      If ADC("UoYN16") = "Y" Then Rep1.DetOdWt.Suppress = False Else Rep1.DetOdWt.Suppress = True
      
      If ADC("UoYN17") = "Y" Then
            Rep1.Subreport1_SecRmCd.Suppress = False
            Rep1.Subreport2_SecRmCd.Suppress = False
            Rep1.Subreport3_SecRmCd.Suppress = False
      Else
            Rep1.Subreport1_SecRmCd.Suppress = True
            Rep1.Subreport2_SecRmCd.Suppress = True
            Rep1.Subreport3_SecRmCd.Suppress = True
      End If
      
      If ADC("UoYN18") = "Y" Then
        Rep1.Subreport1_SecRmSCtg.Suppress = False
        Rep1.Subreport2_SecRmSCtg.Suppress = False
        Rep1.Subreport3_SecRmSCtg.Suppress = False
      Else
        Rep1.Subreport1_SecRmSCtg.Suppress = True
        Rep1.Subreport2_SecRmSCtg.Suppress = True
        Rep1.Subreport3_SecRmSCtg.Suppress = True
      End If
        
        
      If ADC("UoYN19") = "Y" Then
         Rep1.Subreport1_SecRmWt.Suppress = False
         Rep1.Subreport2_SecRmWt.Suppress = False
         Rep1.Subreport3_SecRmWt.Suppress = False
      Else
         Rep1.Subreport1_SecRmWt.Suppress = True
         Rep1.Subreport2_SecRmWt.Suppress = True
         Rep1.Subreport3_SecRmWt.Suppress = True
      End If
         
         
      If ADC("UoYN20") = "Y" Then
        Rep1.Subreport1_SecRmQty.Suppress = False
        Rep1.Subreport2_SecRmQty.Suppress = False
        Rep1.Subreport3_SecRmQty.Suppress = False
      Else
        Rep1.Subreport1_SecRmQty.Suppress = True
        Rep1.Subreport2_SecRmQty.Suppress = True
        Rep1.Subreport3_SecRmQty.Suppress = True
      End If


      If ADC("UoYN21") = "Y" Then
        If Rep1.FldDmCd.Value <> "" Then _
            Rep1.TxtUnitPrcCur.Suppress = False: Rep1.FldOdSalPrc.Suppress = False: Rep1.TxtUnitPrcCurCol.Suppress = False
        
        If Rep1.FldDmCd2.Value <> "" Then _
            Rep1.TxtUnitPrcCur2.Suppress = False: Rep1.FldOdSalPrc2.Suppress = False: Rep1.TxtUnitPrcCurCol2.Suppress = False
        
        If Rep1.FldDmCd3.Value <> "" Then _
            Rep1.TxtUnitPrcCur3.Suppress = False: Rep1.FldOdSalPrc3.Suppress = False: Rep1.TxtUnitPrcCurCol3.Suppress = False
      Else
        Rep1.TxtUnitPrcCur.Suppress = True:  Rep1.FldOdSalPrc.Suppress = True   'Rep1.TxtUnitPrcCurCol.Suppress = True:
        Rep1.TxtUnitPrcCur2.Suppress = True:  Rep1.FldOdSalPrc2.Suppress = True   'Rep1.TxtUnitPrcCurCol.Suppress = True:
        Rep1.TxtUnitPrcCur3.Suppress = True:  Rep1.FldOdSalPrc3.Suppress = True   'Rep1.TxtUnitPrcCurCol3.Suppress = True:
      End If
      
      
      If ADC("UoYN27") = "Y" Then
        
        If Rep1.FldDmCd.Value <> "" Then _
            Rep1.TxtQty.Suppress = False: Rep1.FldqOrdQty.Suppress = False
        
        If Rep1.FldDmCd2.Value <> "" Then _
            Rep1.TxtQty2.Suppress = False: Rep1.FldqOrdQty2.Suppress = False
        
        If Rep1.FldDmCd3.Value <> "" Then _
            Rep1.TxtQty3.Suppress = False: Rep1.FldqOrdQty3.Suppress = False
        
        If ADC("UoYN21") = "N" Then
            Rep1.TxtQty.Left = Rep1.TxtUnitPrcCur.Left: Rep1.FldqOrdQty.Left = Rep1.FldOdSalPrc.Left
            Rep1.TxtQty2.Left = Rep1.TxtUnitPrcCur2.Left: Rep1.FldqOrdQty2.Left = Rep1.FldOdSalPrc2.Left
            Rep1.TxtQty3.Left = Rep1.TxtUnitPrcCur3.Left: Rep1.FldqOrdQty3.Left = Rep1.FldOdSalPrc3.Left
        End If
      Else
        Rep1.TxtQty.Suppress = True: Rep1.FldqOrdQty.Suppress = True
        Rep1.TxtQty2.Suppress = True: Rep1.FldqOrdQty2.Suppress = True
        Rep1.TxtQty3.Suppress = True: Rep1.FldqOrdQty3.Suppress = True
      End If

      If ADC("UoYN21") = "N" And ADC("UoYN27") = "N" Then Rep1.DetOdPrc.Suppress = True

      If ADC("UoYN22") = "Y" Then Rep1.DetodPoNo.Suppress = False Else Rep1.DetodPoNo.Suppress = True
      If ADC("UoYN17") = "N" And ADC("UoYN18") = "N" And ADC("UoYN19") = "N" And ADC("UoYN20") = "N" Then _
      Rep1.DetCatSubRep.Suppress = True Else Rep1.DetCatSubRep.Suppress = False
 

   

    If ADC("UoYN1") = "Y" Then
      Rep.TxtDsgQty.SetText "Amd Qty"
      Rep.TxtDsgQty1.SetText "Amd Qty"
      Rep1.TxtQty.SetText "Amd Qty"
      Rep1.TxtQty2.SetText "Amd Qty"
      Rep1.TxtQty3.SetText "Amd Qty"
    ElseIf ADC("UoYN1") = "N" Then
      Rep.TxtDsgQty.SetText "Ord Qty"
      Rep.TxtDsgQty1.SetText "Ord Qty"
      Rep1.TxtQty.SetText "Ord Qty"
      Rep1.TxtQty2.SetText "Ord Qty"
      Rep1.TxtQty3.SetText "Ord Qty"
    End If

    If ADC("UoYN7") = "Y" Then
      Rep.hBOdDmCd.SetText "Customer Design"
      Rep.Text20.SetText "Customer Design"
      Rep1.Text1.SetText "CmDsg/Sz"
    Else
      Rep.hBOdDmCd.SetText "Design"
      Rep.Text20.SetText "Design"
      Rep1.Text1.SetText "Design/Sz"
    End If

    If ADC("UoYN8") = "Y" Then
        Rep.Text45.SetText "Cust Quality"
      Else
        Rep.Text45.SetText "Quality"
      End If

    If ADC("UoYN9") = "N" And (UCase(ADC.MenuCd) = UCase("RepDetOrd") Or UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") _
                              Or UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg")) Then
      Rep.SecSumm.Suppress = True
      Rep.SecLabSumm.Suppress = True    ' ***** Manali 3.6.0 - 31/10/09 - Lab Summary
    Else
      Rep.SecSumm.Suppress = False
      Rep.SecLabSumm.Suppress = False   ' ***** Manali 3.6.0 - 31/10/09 - Lab Summary
    End If

    If UCase(ADC.MenuCd) <> UCase("RepJtDet") Then
      GRP_REP.SetFormula mRep, "wExpOdNo", ""
      Rep.LblExpOdNo.Suppress = True
    End If

    Rep.TxtOdomPoNo.Suppress = True: Rep.TxtOdPoNo.Suppress = False

    If moCn.GetFldVal("Select HPtrYN from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                      "HCd= '" + ctSelfCmCd + "'") = "N" Then _
       Rep.hOrRmPtr.Suppress = True: Rep.FldOrRmPtr.Suppress = True
    
    CRV_REP.DisplayGroupTree = False
    CmdSC.Top = Me.ScaleHeight - CmdSC.Height
    CmdDispSc.Top = Me.ScaleHeight - CmdDispSc.Height   '***** Manali 3.5.0 - 14/11/08 - Ord Cat fields suppressed/displayed
      
    ' ***** Manali 3.8.0
    If ADC("UoYn25") = "Y" Then
      Rep.FldOdSr.Suppress = True: Rep.FldRunSr.Suppress = False
    Else
      Rep.FldOdSr.Suppress = False: Rep.FldRunSr.Suppress = True
    End If
        
  Call DispReport(ADC) '****** Sachin 3.02 - Display Report ******
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub


Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowPic"
    If .Value = Checked Then ADC("UoYN") = "Y" Else ADC("UoYN") = "N"
  Case Is = 2     'Tag = "ShowDsgVal"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
  Case Is = 4     'Tag = "ShowLabval"
    If .Value = Checked Then ADC("UoYN4") = "Y" Else ADC("UoYN4") = "N"
  Case Is = 6     'Tag = "ShowRem"
    If .Value = Checked Then ADC("UoYN6") = "Y" Else ADC("UoYN6") = "N"
  Case Is = 7     'Tag = "ShowCustDsg"
    If .Value = Checked Then ADC("UoYN7") = "Y" Else ADC("UoYN7") = "N"
  Case Is = 8     'Tag = "ShowCustRm"
    If .Value = Checked Then ADC("UoYN8") = "Y" Else ADC("UoYN8") = "N"
  Case Is = 9     'Tag = "ShowSumm"
    If .Value = Checked Then ADC("UoYN9") = "Y" Else ADC("UoYN9") = "N"
  Case Is = 11     'Tag = "FinalSrt"
    If .Value = Checked Then ADC("UoYN11") = "Y" Else ADC("UoYN11") = "N"
  Case Is = 12     'Tag = "ShowComp"
    If .Value = Checked Then ADC("UoYN12") = "Y" Else ADC("UoYN12") = "N"
  Case Is = 13     'Tag = "ShowGrpDesc"
    If .Value = Checked Then ADC("UoYN13") = "Y" Else ADC("UoYN13") = "N"
  Case Is = 14          'Tag = "ShowRmVal"
    If .Value = Checked Then ADC("UoYN14") = "Y" Else ADC("UoYN14") = "N"
  Case Is = 15          'Tag = "ShowDsgCol"
    If .Value = Checked Then ADC("UoYN15") = "Y" Else ADC("UoYN15") = "N"
  Case Is = 16         'Tag = "ShowGrsWt"
    If .Value = Checked Then ADC("UoYN16") = "Y" Else ADC("UoYN16") = "N"
  Case Is = 17          'Tag = "ShowRmCd"
    If .Value = Checked Then ADC("UoYN17") = "Y" Else ADC("UoYN17") = "N"
  Case Is = 18          'Tag = "ShowRmSCtg"
    If .Value = Checked Then ADC("UoYN18") = "Y" Else ADC("UoYN18") = "N"
  Case Is = 19          'Tag = "ShowRmWt"
    If .Value = Checked Then ADC("UoYN19") = "Y" Else ADC("UoYN19") = "N"
  Case Is = 20         'Tag = "ShowRmQty"
    If .Value = Checked Then ADC("UoYN20") = "Y" Else ADC("UoYN20") = "N"
  Case Is = 21         'Tag = "ShowUnitPrc"
    If .Value = Checked Then ADC("UoYN21") = "Y" Else ADC("UoYN21") = "N"
  Case Is = 22          'Tag = "ShowOdPO"
    If .Value = Checked Then ADC("UoYN22") = "Y" Else ADC("UoYN22") = "N"
  Case Is = 24          'Tag = "ShowVchSr"
    If .Value = Checked Then ADC("UoYN24") = "Y" Else ADC("UoYN24") = "N"
  Case Is = 25          'Tag = "ShowVchSr"    ' ***** Manali 3.8.0
    If .Value = Checked Then ADC("UoYN25") = "Y" Else ADC("UoYN25") = "N"
  Case Is = 26          'Tag = "ShowMstRem"    ' ***** Manali 3.9.0
    If .Value = Checked Then ADC("UoYN26") = "Y" Else ADC("UoYN26") = "N"
  Case Is = 27          'Tag = "ShowOrdQty"
    If .Value = Checked Then ADC("UoYN27") = "Y" Else ADC("UoYN27") = "N"
  Case Is = 28     'Tag = "ShowDistDsg"
    If .Value = Checked Then ADC("UoYN28") = "Y" Else ADC("UoYN28") = "N"       '4.1.2
End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
'Private Sub CmdChgImg_Click()
'If FraChgImg.Visible = False Then
'    If CRV_REP.Visible = True Then Exit Sub
'    FraChgImg.Enabled = True
'    FraChgImg.Visible = True
'    TAB_REP.Enabled = False
'    adc.AllowSave = False
'    FraChgImg.ZOrder
'  Else
'    TAB_REP.Enabled = True
'    adc.AllowSave = True
'    FraChgImg.Enabled = False
'    FraChgImg.Visible = False
'  End If
'End Sub
'Private Sub CmdChgImgGo_Click()
'On Error GoTo errhd
'Dim wPicHeight As Single, wPicWidth As Single
'If UCase(adc.MenuCd) = UCase("RepOrdCat") Or UCase(adc.MenuCd) = UCase("RepFgCat") Or UCase(adc.MenuCd) = UCase("RepJtCat") Then
'    wPicHeight = 110: wPicWidth = 110
'Else
'    wPicHeight = 84: wPicWidth = 84
'End If
'
'If adc("wXlFilePath") = "" Then
'    DispMsg "Please Provide an Excel File to change Images", etError
'    Exit Sub
'End If
'
'If ReplXlImg(adc("wXlFilePath"), wo_rsPic, wPicHeight, wPicWidth) Then
'    TAB_REP.Enabled = True
'    adc.AllowSave = True
'    FraChgImg.Visible = False
'    CmdChgImg.Visible = False
'    DispMsg "Images Replaced Successfully", etInfo
'End If
'  Exit Sub
'errhd:
'    DispMsg Err.Description, etError
'End Sub
Private Sub CmdDispSc_Click()
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
  If FraDispSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraDispSc.Enabled = True
    FraDispSc.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraDispSc.ZOrder
    'adc("UoYN15").SetFocus
    CmdSC.Enabled = False
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraDispSc.Enabled = False
    FraDispSc.Visible = False
    CmdDispSc.SetFocus
    CmdSC.Enabled = True
  End If
' **** Manali 3.5.0 - 14/11/08 - Supress or Display fields in Ord Cat
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
Private Sub ADC_Load()
    Set moCn = ADC.Connection    '*** (Bef speed) bef 24/11/06
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
        UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      Set ADC.FirNKeyCtl = ADC("UoOmTcFr") 'adc("UoCoCdFr")
    Else
      Set ADC.FirNKeyCtl = ADC("UoTcTypFr") 'adc("UoCoCdFr")
    End If
    '*** Jay 3.3(FG)
    Call InitProp(Me)
    
    ' ***** Manali Trading Module
    gs_CmCtg = "C"
    gs_EntTyp = "FG"
      ADC("UoYN").Visible = False: ADC("UoYN").Enabled = False
      ChkBoxArr(0).Enabled = False: ChkBoxArr(0).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      LblPic.Visible = False
      ADC("UoYN2").Visible = False: ADC("UoYN2").Enabled = False: LblShwVal.Visible = False
      ChkBoxArr(2).Enabled = False: ChkBoxArr(2).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ADC("UoYN4").Visible = False: ADC("UoYN4").Enabled = False: LblLabVal.Visible = False
      ChkBoxArr(4).Enabled = False: ChkBoxArr(4).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** Jay 2.14Next [Show Rm Value]
      ADC("UoYN14").Visible = False: ADC("UoYN14").Enabled = False: LblRmVal.Visible = False
      ChkBoxArr(14).Enabled = False: ChkBoxArr(14).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      
      ' urmi show remark
      ADC("UoYN6").Visible = False: ADC("UoYN6").Enabled = False: lblShowRem.Visible = False
      ChkBoxArr(6).Enabled = False: ChkBoxArr(6).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ChkBoxArr(26).Enabled = False: ChkBoxArr(26).Visible = False    ' ****** Manali 3.9.0 - Different Option For Design level Remark
      ' Manoj 2.10.0
      ADC("UoYN9").Visible = False: ADC("UoYN9").Enabled = False: LblShowSumm.Visible = False
      ChkBoxArr(9).Enabled = False: ChkBoxArr(9).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      ADC("UoYn10").Visible = False: ADC("UoYN10").Enabled = False: LblCalcPtr.Visible = False
      ' Manoj 2.10.0
      ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      FraOptPtr.Visible = False: FraOptPtr.Visible = False
       ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
      '*** JAY 2.14 ***
      ADC("UOYN13").Visible = False: LblShowGrpDesc.Visible = False
      ChkBoxArr(13).Enabled = False: ChkBoxArr(13).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
      '*** JAY 2.14 ***
      
      ' ***** Manali 3.8.0
      ChkBoxArr(25).Enabled = False: ChkBoxArr(25).Visible = False
       
     
        'gs_OmTcTyp = "FB"
        gs_EntTyp = ADC("UOTCTYPFR")
        
       lblAmdQty.Visible = False:  ADC("UOYn1").Visible = False
       FraOptOrdQty.Enabled = False: FraOptOrdQty.Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
       lblPrdRmQtyWt.Visible = False: ADC("UOYn3").Visible = False
       LblCalcPtr.Visible = False: ADC("UOYn10").Visible = False
       ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons added
       FraOptRm.Enabled = False: FraOptRm.Visible = False
       FraOptPtr.Visible = False: FraOptPtr.Visible = False
       FraOptOrdQty.Enabled = False: FraOptOrdQty.Visible = False
    
    ' ***** Manali 3.6.0 - Order Printing in Excel Format Module
    If moCn.RecSeek("Select 'x' from Head Where HCoCd='" + ctSelfCoCd + "' and HCd='" + ctSelfCmCd + "' and HOrdExcelRepYn<>'Y' ") Then _
      ChkBoxArr(23).Visible = False: ChkBoxArr(23).Enabled = False
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
   '*** Jay 3.3(FG)
     gs_FgTcTyp = "FB"
     gs_EntTyp = "FGBAG" 'adc("UOTCTYPFR")
     gs_TxnTcTyp = "FM"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  gs_TxnTcTyp = "FM"
   
  Select Case UCase(IdName)
  Case Is = UCase("UOTCTYPFR")
    gs_EntTyp = "FGBAG"
  Case Is = UCase("UoFgTcFr"), UCase("UoFgTcTo")
    gs_FgTcTyp = "FB"
      
    '  gs_OmTcTyp = adc("UoTcTypFr")
     '  gs_PTyp = "PCT"
     ' gs_Tbl = "Param"
    
  Case Is = UCase("UoYN4")
    
  Case Is = UCase("UoYN10")
    Cancel = IIF(ADC("UoYN3") = "N", True, False)
    ErrMsg = "Can be Used Only When Prd Wt & Prd Qty is Used": Exit Sub
  End Select
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  Select Case UCase(IdName)
  Case Is = UCase("UoYN10")
    If ADC("UoYN3") = "N" Then ADC("UoYN10") = "N"
  End Select
End Sub
Private Sub CmdSc_Click()
  If FraSC.Visible = False Then
    FraSC.Enabled = True
    FraSC.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    CmdDispSc.Enabled = False   ' ******** Manali 3.5.0 - 14/11/08
    FraSC.ZOrder
  Else
    TAB_REP.Enabled = True
    FraSC.Enabled = False
    FraSC.Visible = False
    ADC.AllowSave = True
    CmdSC.SetFocus
    CmdDispSc.Enabled = True   ' ******** Manali 3.5.0 - 14/11/08
  End If
End Sub

Private Sub OptOrdQty_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN1") = "Y"
Case Is = 1
  ADC("UoYN1") = "N"
End Select
End Sub
Private Sub OptOrdQty_GotFocus(Index As Integer)
  DispMsg FraOptOrdQty.ToolTipText, etInfo
End Sub
Private Sub OptPtr_Click(Index As Integer)
Select Case Index
Case Is = 0
  If OptRmPrd(1).Value = True Then
    OptPtr(0).Value = False: OptPtr(1).Value = True: ADC("UoYN10") = "N"
    DispMsg "Calculated Pointer can be selected only when Prd Wt & Qty is Used", etError
  Else
    ADC("UoYN10") = "Y"
  End If
Case Is = 1
  ADC("UoYN10") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptPtr_GotFocus(Index As Integer)
  DispMsg FraOptPtr.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptRmPrd_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN3") = "Y"
Case Is = 1
  ADC("UoYN3") = "N"
  OptPtr(1).Value = True: ADC("UoYN10") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
Private Sub OptRmPrd_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    ADC("UOUSRCD") = ADC.UsrCd
    ADC("UOMNUCD") = ADC.MenuCd
     '*************************** ZUBIN **************************
  ' 08th Dec 2003, EMR206
  '*** Jay 3.3(FG)
  ' **** Manali Trading Module
  If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
    UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
  'If UCase(ADC.MenuCd) = UCase("RepAvgOrd") Or UCase(ADC.MenuCd) = UCase("RepOrdCat") Then
  '*** Jay 3.3(FG)
    lblCustRm.Visible = False
    ADC("UoYN8").Visible = False
    ChkBoxArr(8).Enabled = False: ChkBoxArr(8).Visible = False    ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  End If
  '*************************** ZUBIN **************************
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
    '*************Geeta*********************
    If UCase(ADC("UoTcTypFr")) = "PL" Then
      ADC("UoCoCdFr") = ctSelfCoCd
    Else
      ADC("UoCoCdFr") = gs_CoCd
    End If
    '*** Jay 3.3(FG)
    If UCase(ADC.MenuCd) = UCase("RepFgDet") Or UCase(ADC.MenuCd) = UCase("RepFgAvg") Or UCase(ADC.MenuCd) = UCase("RepFgCat") Or _
      UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat") Then
      ADC("UODELDTFR") = MWLib.EmptyDate: ADC("UODELDTTo") = MWLib.EmptyDate
      ADC("UODELDT1FR") = MWLib.EmptyDate: ADC("UODELDT1To") = MWLib.EmptyDate
      ADC("UOYn1") = "Y": ADC("UOYn3") = "N": ADC("UOYn10") = "N"
      ADC("UoTcTypFr") = IIF(UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Or UCase(ADC.MenuCd) = UCase("RepJtCat"), ctJtTcTyp, "FG")
    End If
    '*** Jay 3.3(FG)
 '*******************************
  ' adc("UoCoCdFr") = gs_CoCd
  ' Zubin 212
  '*** Jay 3.3(FG)
  If ADC("UOYN4") = "" And (UCase(ADC.MenuCd) <> UCase("RepOrdCat") And UCase(ADC.MenuCd) <> UCase("RepFgCat") And UCase(ADC.MenuCd) <> UCase("RepJtCat")) Then ADC("UOYN4") = "Y"
  'If ADC("UOYN4") = "" And UCase(ADC.MenuCd) <> UCase("RepOrdCat") Then ADC("UOYN4") = "Y"
  '*** Jay 3.3(FG)
  ' Zubin 212
  
  If ADC("UoYN") = "Y" Then
    OptBag(0).Value = True
  ElseIf ADC("UoYN") = "N" Then
    OptBag(1).Value = True
  ElseIf ADC("UoYN") = "" Then
    OptBag(2).Value = True
  End If
  
  ' ****** Manali 3.5.0 - 07/11/08 - Radio Buttons Added
  If ADC("UoYN5") = "Y" Then OptDelDt(0).Value = True
  If ADC("UoYN5") = "N" Then OptDelDt(1).Value = True
  If ADC("UoYN3") = "Y" Then OptRmPrd(0).Value = True
  If ADC("UoYN3") = "N" Then OptRmPrd(1).Value = True
  If ADC("UoYN10") = "Y" Then OptPtr(0).Value = True
  If ADC("UoYN10") = "N" Then OptPtr(1).Value = True
  If ADC("UoYN1") = "Y" Then OptOrdQty(0).Value = True
  If ADC("UoYN1") = "N" Then OptOrdQty(1).Value = True
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If ADC("UoYN") = "" Then ADC("UoYN") = "N"
  If ADC("UoYN2") = "" Then ADC("UoYN2") = "N"
  'If adc("UoYN4") = "" Then adc("UoYN4") = "N"
  If ADC("UoYN6") = "" Then ADC("UoYN6") = "N"
  If ADC("UoYN7") = "" Then ADC("UoYN7") = "N"
  If ADC("UoYN8") = "" Then ADC("UoYN8") = "N"
  If ADC("UoYN9") = "" Then ADC("UoYN9") = "N"
  If ADC("UoYN11") = "" Then ADC("UoYN11") = "N"
  If ADC("UoYN13") = "" Then ADC("UoYN13") = "N"
  If ADC("UoYN14") = "" Then ADC("UoYN14") = "N"
  If ADC("UoYN15") = "" Then ADC("UoYN15") = "N"
  If ADC("UoYN16") = "" Then ADC("UoYN16") = "N"
  If ADC("UoYN17") = "" Then ADC("UoYN17") = "N"
  If ADC("UoYN18") = "" Then ADC("UoYN18") = "N"
  If ADC("UoYN19") = "" Then ADC("UoYN19") = "N"
  If ADC("UoYN20") = "" Then ADC("UoYN20") = "N"
  If ADC("UoYN21") = "" Then ADC("UoYN21") = "N"
  If ADC("UoYN22") = "" Then ADC("UoYN22") = "N"
  If ADC("UoYN24") = "" Then ADC("UoYN24") = "N"
  If ADC("UoYN25") = "" Then ADC("UoYN25") = "N"      ''''' Manali 3.8.0
  If ADC("UoYN26") = "" Then ADC("UoYN26") = "N"      ''''' Manali 3.9.0 Design Level Remarks
  If ADC("UoYN27") = "" Then ADC("UoYN27") = "N"      ' ***** Manali 3.10.0 - 09/03/12 - Different Suppres option for Unit Pricce and Qty
  If ADC("UoYN28") = "" Then ADC("UoYN28") = "N"      '4.1.2

  If ADC("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked
  If ADC("UoYN4") = "Y" Then ChkBoxArr(4).Value = Checked Else ChkBoxArr(4).Value = Unchecked
  If ADC("UoYN6") = "Y" Then ChkBoxArr(6).Value = Checked Else ChkBoxArr(6).Value = Unchecked
  If ADC("UoYN7") = "Y" Then ChkBoxArr(7).Value = Checked Else ChkBoxArr(7).Value = Unchecked
  If ADC("UoYN8") = "Y" Then ChkBoxArr(8).Value = Checked Else ChkBoxArr(8).Value = Unchecked
  If ADC("UoYN9") = "Y" Then ChkBoxArr(9).Value = Checked Else ChkBoxArr(9).Value = Unchecked
  If ADC("UoYN11") = "Y" Then ChkBoxArr(11).Value = Checked Else ChkBoxArr(11).Value = Unchecked
  If ADC("UoYN13") = "Y" Then ChkBoxArr(13).Value = Checked Else ChkBoxArr(13).Value = Unchecked
  If ADC("UoYN14") = "Y" Then ChkBoxArr(14).Value = Checked Else ChkBoxArr(14).Value = Unchecked
  If ADC("UoYN15") = "Y" Then ChkBoxArr(15).Value = Checked Else ChkBoxArr(15).Value = Unchecked
  If ADC("UoYN16") = "Y" Then ChkBoxArr(16).Value = Checked Else ChkBoxArr(16).Value = Unchecked
  If ADC("UoYN17") = "Y" Then ChkBoxArr(17).Value = Checked Else ChkBoxArr(17).Value = Unchecked
  If ADC("UoYN18") = "Y" Then ChkBoxArr(18).Value = Checked Else ChkBoxArr(18).Value = Unchecked
  If ADC("UoYN19") = "Y" Then ChkBoxArr(19).Value = Checked Else ChkBoxArr(19).Value = Unchecked
  If ADC("UoYN20") = "Y" Then ChkBoxArr(20).Value = Checked Else ChkBoxArr(20).Value = Unchecked
  If ADC("UoYN21") = "Y" Then ChkBoxArr(21).Value = Checked Else ChkBoxArr(21).Value = Unchecked
  If ADC("UoYN22") = "Y" Then ChkBoxArr(22).Value = Checked Else ChkBoxArr(22).Value = Unchecked
  If ADC("UoYN24") = "Y" Then ChkBoxArr(24).Value = Checked Else ChkBoxArr(24).Value = Unchecked
  If ADC("UoYN25") = "Y" Then ChkBoxArr(25).Value = Checked Else ChkBoxArr(25).Value = Unchecked    ''''' Manali 3.8.0
  If ADC("UoYN26") = "Y" Then ChkBoxArr(26).Value = Checked Else ChkBoxArr(26).Value = Unchecked    ' ***** Manali 3.9.0 - Design Level Remarks
  If ADC("UoYN27") = "Y" Then ChkBoxArr(27).Value = Checked Else ChkBoxArr(27).Value = Unchecked    ' ***** Manali 3.10.0 - 09/03/12 - Different Suppres option for Unit Pricce and Qty
  If ADC("UoYN28") = "Y" Then ChkBoxArr(28).Value = Checked Else ChkBoxArr(28).Value = Unchecked    '4.1.2
  
End Sub
Private Sub DispCoNm()
'mRep.wCoCd.SetText gs_CoNm
    Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + ADC("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    
    If ADC("UoTcTypFr") = "PL" Then
      mRep.wCoCd.SetText gs_CoNm
      mRep.wCoCdLogo.SetText gs_CoNm
    Else
      mRep.wCoCd.SetText ws_HName
      mRep.wCoCdLogo.SetText ws_HName
    End If
  
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: mRep.wCoCd.Suppress = True
    If UCase(ADC.MenuCd) = UCase("RepJtDet") Or UCase(ADC.MenuCd) = UCase("RepJtAvg") Then
      Rep.RHLogo.UnderlaySection = True
    End If
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ********* Manali 3.2.2 - Before 3.2.2
'  Dim ws_DtTm As String
'  ws_DtTm = Format$(moCn.SrvrDate, "DD-MMM-YYYY")
'  mRep.wSrvrDate.SetText ws_DtTm
'  ws_DtTm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
'  mRep.wSrvrTime.SetText ws_DtTm

' ******** Manali 3.2.2 - User date and time in reports
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
    mRep.TxtRepUsrDtTime.SetText Trim(ADC.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

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
Private Sub ATXT_Validate(Index As Integer, Cancel As Boolean)
    Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    '*** (09/08/05)
    Set moCn = Nothing   '*** (Bef speed) bef 24/11/06
    Set mRep = Nothing
    Set Rep = Nothing
    Set Rep1 = Nothing
    '*** (09/08/05)
End Sub
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  ADC("UoYN") = "Y"
Case Is = 1
  ADC("UoYN") = "N"
Case Is = 2
  ADC("UoYN") = ""
End Select
End Sub
Private Sub OptBag_GotFocus(Index As Integer)
  DispMsg FraOptBag.ToolTipText, etInfo
End Sub


Private Function MakeStr(wsPassStr As String) As String
Dim i As Integer, NewStr As String
  
  i = 1
  Do While i <= Len(wsPassStr)
    If Mid(wsPassStr, i, 1) <> Chr(13) Then
      NewStr = NewStr + Mid(wsPassStr, i, 1)
    Else
      NewStr = NewStr + " "
    End If
    i = i + 1
  Loop
  MakeStr = NewStr
End Function

