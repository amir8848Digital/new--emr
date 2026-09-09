VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpJtStkRm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Rm"
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
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2280
      TabIndex        =   96
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   688
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "Det S&cope"
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
      Left            =   11640
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4410
      TabIndex        =   63
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
      Left            =   5745
      TabIndex        =   69
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   70
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
         TabIndex        =   71
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
      Left            =   6960
      TabIndex        =   68
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   67
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
      TabIndex        =   64
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
      Height          =   9675
      Left            =   0
      TabIndex        =   66
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   60
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpJtStkRm.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpJtStkRm.frx":001C
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
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Include Sec SalCtg"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   460
               Index           =   0
               Left            =   60
               TabIndex        =   18
               Tag             =   "SalCtg2,3"
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   2250
               Width           =   2265
            End
            Begin VB.Frame FraOptValWt 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1900
               TabIndex        =   102
               ToolTipText     =   "Select if to show Value or Weight"
               Top             =   6735
               Width           =   2940
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Value"
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
                  Index           =   0
                  Left            =   180
                  TabIndex        =   104
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.OptionButton OptValWt 
                  Caption         =   "Weight"
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
                  Index           =   1
                  Left            =   1605
                  TabIndex        =   103
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4800
               TabIndex        =   13
               ToolTipText     =   "Enter To Design Category"
               Top             =   1680
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
               Index           =   0
               Left            =   7500
               TabIndex        =   14
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1680
               Width           =   7410
               _ExtentX        =   13070
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
               Left            =   2115
               TabIndex        =   12
               ToolTipText     =   "Enter From Design Category"
               Top             =   1680
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
               Left            =   4800
               TabIndex        =   20
               ToolTipText     =   "Enter To Design Code"
               Top             =   2730
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
               CmpStr          =   "DDmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   7500
               TabIndex        =   21
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2730
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "DDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   2115
               TabIndex        =   19
               ToolTipText     =   "Enter From Design Code"
               Top             =   2730
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
               CmpStr          =   "DDmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4800
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Sales Category"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGTO"
               IdName          =   "UODMSALCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   7500
               TabIndex        =   17
               ToolTipText     =   "Enter Design Sales Category Selection"
               Top             =   1965
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMSALCTGSEL"
               IdName          =   "UODMSALCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2115
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Sales Category"
               Top             =   1965
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMSALCTGFR"
               IdName          =   "UODMSALCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4800
               TabIndex        =   23
               ToolTipText     =   "Enter To Design Size"
               Top             =   3015
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
               CmpStr          =   "DDmSz <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   2115
               TabIndex        =   22
               ToolTipText     =   "Enter From Design Size"
               Top             =   3015
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
               CmpStr          =   "DDmSz >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   2115
               TabIndex        =   30
               ToolTipText     =   "Enter From Suffix"
               Top             =   3870
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
               CmpStr          =   "DOdSfx >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2115
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "DCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   73
               Left            =   4800
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   540
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "DCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   75
               Left            =   7500
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   540
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "DCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4800
               TabIndex        =   50
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   5295
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
               Index           =   38
               Left            =   7500
               TabIndex        =   51
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   5295
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "OrRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   4800
               TabIndex        =   47
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   5010
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
               Index           =   46
               Left            =   7500
               TabIndex        =   48
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   5010
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "OrRmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   2115
               TabIndex        =   46
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   5010
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
               Index           =   51
               Left            =   2115
               TabIndex        =   49
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   5295
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
               Index           =   81
               Left            =   7500
               TabIndex        =   42
               ToolTipText     =   "Enter Raw Material Category Selection"
               Top             =   4440
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCTGSEL"
               IdName          =   "UORMCTGSEL"
               CmpStr          =   "OrRmCtg In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4800
               TabIndex        =   10
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1395
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "DLoc<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   7500
               TabIndex        =   11
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1395
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "DLoc in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   2115
               TabIndex        =   9
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1395
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "DLoc>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   2115
               TabIndex        =   3
               ToolTipText     =   "Enter Location Type"
               Top             =   825
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPFR"
               IdName          =   "UOLOCTYPFR"
               CmpStr          =   "LocTyp>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   27
               Left            =   4800
               TabIndex        =   4
               ToolTipText     =   "Enter To Location Type"
               Top             =   825
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOLOCTYPTO"
               IdName          =   "UOLOCTYPTO"
               CmpStr          =   "LocTyp<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   7500
               TabIndex        =   5
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   825
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLOCTYPSEL"
               IdName          =   "UOLOCTYPSEL"
               CmpStr          =   "LocTyp  in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   4800
               TabIndex        =   7
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1110
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
               Index           =   7
               Left            =   7500
               TabIndex        =   8
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1110
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   2115
               TabIndex        =   6
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1110
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
               Index           =   9
               Left            =   4800
               TabIndex        =   44
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   4725
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSTO"
               IdName          =   "UOSTNCLSTO"
               CmpStr          =   "PSCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   7500
               TabIndex        =   45
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   4725
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "PSCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4800
               TabIndex        =   31
               ToolTipText     =   "Enter To Suffix"
               Top             =   3870
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   10
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
               CmpStr          =   "DOdSfx <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   79
               Left            =   4800
               TabIndex        =   41
               ToolTipText     =   "Enter To Raw Material Category"
               Top             =   4440
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGTO"
               IdName          =   "UORMCTGTO"
               CmpStr          =   "OrRmCtg <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   2115
               TabIndex        =   56
               ToolTipText     =   "Enter From YearMonth"
               Top             =   6150
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               DataField       =   "UOYYMMFR"
               ReCalcOn        =   "UODTFR"
               IdName          =   "UOYYMMFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   4800
               TabIndex        =   53
               ToolTipText     =   "Enter To Raw Material Len/Sv"
               Top             =   5580
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
               Index           =   54
               Left            =   2115
               TabIndex        =   52
               ToolTipText     =   "Enter From Raw Material Len/Sv"
               Top             =   5580
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
               Index           =   56
               Left            =   4800
               TabIndex        =   55
               ToolTipText     =   "Enter To Raw Material Breadth"
               Top             =   5865
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
               Index           =   57
               Left            =   2115
               TabIndex        =   54
               ToolTipText     =   "Enter From Raw Material Breadth"
               Top             =   5865
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
               Index           =   16
               Left            =   3855
               TabIndex        =   35
               ToolTipText     =   "Enter From Stock Number"
               Top             =   4155
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDFR"
               IdName          =   "UOODDMIDFR"
               CmpStr          =   "DDmStkNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   6480
               TabIndex        =   39
               ToolTipText     =   "Enter To Stock Number"
               Top             =   4155
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDTO"
               IdName          =   "UOODDMIDTO"
               CmpStr          =   "DDmStkNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   3120
               TabIndex        =   34
               ToolTipText     =   "Enter From Stock Character"
               Top             =   4155
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "DDmStkChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2115
               TabIndex        =   43
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   4725
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOSTNCLSFR"
               IdName          =   "UOSTNCLSFR"
               CmpStr          =   "PSCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   2115
               TabIndex        =   40
               ToolTipText     =   "Enter From Raw Material Category"
               Top             =   4440
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "OrRmCtg >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2715
               TabIndex        =   33
               ToolTipText     =   "Enter From Stock Year"
               Top             =   4155
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "DDmStkYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   2115
               TabIndex        =   57
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   6450
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               ReCalcParent    =   "UOYYMMFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   2115
               TabIndex        =   32
               ToolTipText     =   "Enter From Stock company code"
               Top             =   4155
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDFR"
               IdName          =   "UOSTKCOCDFR"
               CmpStr          =   "DdmStkCoCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   5805
               TabIndex        =   38
               ToolTipText     =   "Enter To Stock Character"
               Top             =   4155
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "DDmStkChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   22
               Left            =   5400
               TabIndex        =   37
               ToolTipText     =   "Enter To Stock Year"
               Top             =   4155
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "DDmStkYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   4800
               TabIndex        =   36
               ToolTipText     =   "Enter To Stock Company Code"
               Top             =   4155
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDTO"
               IdName          =   "UOSTKCOCDTO"
               CmpStr          =   "DdmStkCoCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   4800
               TabIndex        =   25
               ToolTipText     =   "Enter To Design Colour"
               Top             =   3300
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLTO"
               IdName          =   "UODMCOLTO"
               CmpStr          =   "Odm.OdDmCol <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   7500
               TabIndex        =   26
               ToolTipText     =   "Enter Design Colour Selection"
               Top             =   3300
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCOLSEL"
               IdName          =   "UODMCOLSEL"
               CmpStr          =   "Odm.OdDmCol In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   2115
               TabIndex        =   24
               ToolTipText     =   "Enter From Design Colour"
               Top             =   3300
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UODMCOLFR"
               IdName          =   "UODMCOLFR"
               CmpStr          =   "Odm.OdDmCol >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4800
               TabIndex        =   28
               ToolTipText     =   "Enter To Karat"
               Top             =   3585
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTTO"
               IdName          =   "UOKTTO"
               CmpStr          =   "Odm.OdKt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   7500
               TabIndex        =   29
               ToolTipText     =   "Enter Karat Selection"
               Top             =   3585
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOKTSEL"
               IdName          =   "UOKTSEL"
               CmpStr          =   "Odm.OdKt In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   2115
               TabIndex        =   27
               ToolTipText     =   "Enter From Karat"
               Top             =   3585
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOKTFR"
               IdName          =   "UOKTFR"
               CmpStr          =   "Odm.OdKt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   4995
               TabIndex        =   105
               ToolTipText     =   "Show Open Or Closed Or All Bags ? (Open - 'Y', Closed - 'N', All - Blank)"
               Top             =   6735
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   3840
               TabIndex        =   106
               ToolTipText     =   "Check To Include SalCtg2 and SalCtg3"
               Top             =   2280
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WADDALLSALCTG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   2115
               TabIndex        =   58
               ToolTipText     =   "Enter Divided By Amount "
               Top             =   7080
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               Alignment       =   2
               Mask            =   "#####0"
               MaxLength       =   6
               DataType        =   1
               IdName          =   "UOVALDIVBY"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2115
               TabIndex        =   59
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   7365
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UOCURCDFR"
               IdName          =   "UOCURCDFR"
            End
            Begin VB.Label LblCurCd 
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
               TabIndex        =   108
               Top             =   7365
               Width           =   1185
            End
            Begin VB.Label LblValDivBy 
               BackStyle       =   0  'Transparent
               Caption         =   "Value Divided By"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   107
               Top             =   7080
               Width           =   1665
            End
            Begin VB.Label LblVal 
               BackStyle       =   0  'Transparent
               Caption         =   "Show "
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   101
               Top             =   6735
               Width           =   1665
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
               Index           =   9
               Left            =   90
               TabIndex        =   100
               Top             =   3585
               Width           =   1365
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Design Colour"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   99
               Top             =   3300
               Width           =   1395
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "If As Of Date <> 01/01/80 report will be very slow. NOT to be used during working hours"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   3360
               TabIndex        =   95
               Top             =   6450
               Width           =   9105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Set As Of Date = 01/01/80 to edit YyyyMm"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   3360
               TabIndex        =   94
               Top             =   6150
               Width           =   4575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "As of Date"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   60
               TabIndex        =   93
               Top             =   6450
               Width           =   1425
            End
            Begin VB.Label LblDmId 
               BackStyle       =   0  'Transparent
               Caption         =   "StkNum"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   92
               Top             =   4155
               Width           =   1425
            End
            Begin VB.Label LblLen 
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
               Left            =   90
               TabIndex        =   91
               Top             =   5580
               Width           =   1425
            End
            Begin VB.Label LblBrdth 
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
               Left            =   90
               TabIndex        =   90
               Top             =   5865
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "YyyyMm"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   90
               TabIndex        =   89
               Top             =   6150
               Width           =   1425
            End
            Begin VB.Label LblStnCls 
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
               Left            =   90
               TabIndex        =   88
               Top             =   4725
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
               Left            =   90
               TabIndex        =   87
               Top             =   1425
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
               Left            =   90
               TabIndex        =   86
               Top             =   825
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
               Left            =   90
               TabIndex        =   85
               Top             =   1110
               Width           =   1425
            End
            Begin VB.Label LblRmCd 
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
               Left            =   90
               TabIndex        =   84
               Top             =   5295
               Width           =   1425
            End
            Begin VB.Label LblRmSubCtg 
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
               Left            =   90
               TabIndex        =   83
               Top             =   5010
               Width           =   1425
            End
            Begin VB.Label LblRmCtg 
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
               Left            =   90
               TabIndex        =   82
               ToolTipText     =   "Location"
               Top             =   4440
               Width           =   1425
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
               Left            =   4800
               TabIndex        =   81
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
               Left            =   2115
               TabIndex        =   80
               Top             =   60
               Width           =   1335
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
               Index           =   6
               Left            =   90
               TabIndex        =   79
               Top             =   1965
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
               TabIndex        =   78
               Top             =   1680
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
               Left            =   90
               TabIndex        =   77
               Top             =   2730
               Width           =   1425
            End
            Begin VB.Label LblDmSz 
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
               Left            =   90
               TabIndex        =   76
               Top             =   3015
               Width           =   1365
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
               Left            =   90
               TabIndex        =   75
               Top             =   3870
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
               Index           =   20
               Left            =   7500
               TabIndex        =   74
               Top             =   60
               Width           =   1575
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
               Index           =   21
               Left            =   90
               TabIndex        =   73
               Top             =   540
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
               TabIndex        =   62
               Top             =   360
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   635
            End
         End
      End
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1320
         TabIndex        =   97
         Top             =   6480
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   4286
            Cols            =   4
            colname1        =   "UDANASR"
            heading1        =   "Sr. "
            datafld1        =   "UdAnaSr"
            datatype1       =   4
            recalcparent1   =   "GrdDsgAna(UDANACDFR),GrdDsgAna(UDANACDTO)"
            maxlength1      =   2
            tooltiptext1    =   "Enter Analysis Sr. Number"
            colname2        =   "UDANACDFR"
            heading2        =   "Fr Ana Cd      "
            datafld2        =   "UdAnaCdFr"
            datatype2       =   4
            recalcon2       =   "GrdDsgAna(UDANASR)"
            maxlength2      =   8
            tooltiptext2    =   "Enter From Analysis Code"
            colname3        =   "UDANACDTO"
            heading3        =   "To Ana Cd      "
            datafld3        =   "UdAnaCdTo"
            datatype3       =   4
            recalcon3       =   "GrdDsgAna(UDANASR)"
            maxlength3      =   8
            tooltiptext3    =   "Enter To Analysis Code"
         End
      End
   End
End
Attribute VB_Name = "EmrFrpJtStkRm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRep As CRAXDRT.Report
Dim Rep As New EmrRepJtStkRm
Dim Rep1 As New EmrRepJtDsgFlashStk
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim ws_MultiCoMod As Boolean
Dim ws_hAddJROYN As String
Private Sub SetGroupSort()
  If UCase(adc.MenuCd) = UCase("RepJtStkRm") Then
   GRP_REP.Add "Loc Typ", "LocTyp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp = 'LOCTYP' And PMCD = LocTyp) "
  GRP_REP.Add "(Typ) Loc", "RTrim(LocTyp)+RTrim(DCoCd)+RTrim(DLoc)", "'(' + RTrim(LocTyp) + '/'+RTrim(DCoCd)+') '+ RTrim(DLoc)", "", "", "Company Code", "", "", " LocDesc"
  
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
  'uni.4 DmSalCtg replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=DmCtg and PSCd= ||qDmSalCtgDescRepl|| ) "
  GRP_REP.Add "Design Code", "DDmCd", "", "", "", "Design Ctg,Sales Ctg", "", "", "DmDesc "
  GRP_REP.Add "Design Size", "DDmSz", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= DDmSz)"
  GRP_REP.Add "Suffix", "DOdSfx", "", "", "", "", "", "", ""
  'sv.48 DBfldname copied to DispFldNm and DBfldname format changed
  GRP_REP.Add "DmStkNum", "DdmStkCoCd+DDmStkYy+DDmStkChr+Str(DDmStkNo)", "DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+Ltrim(Str(DDmStkNo))", "", "", "", "", "", ""
  'sv.48 In DBfldname ltrim removed from DDmStkNo
  GRP_REP.Add "DsgCd-Sfx-Sz-StkNum", "DDmCd+DOdSfx+DDmSz+DdmStkCoCd+DDmStkYy+DDmStkChr+Str(DDmStkNo)", "DDmCd+'/'+DOdSfx+'/'+DDmSz+'/'+DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+LTrim(Str(DDmStkNo))", "", "", "Design Ctg,Sales Ctg,Design Code,Design Size,Suffix,DmStkNum", "", "", ""

  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "DCoCd", "DCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=DCoCd) "

  GRP_REP.Add "Rm Ctg", "OrRmCtg", "", "OrRmCtg", "hOrRmCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMCTG' and PMCd = OrRmCtg) "
  GRP_REP.Add "Rm SubCtg", "OrRmCtg+OrRmSCtg", "'('+ RTrim(OrRmCtg)+ ') '+ OrRmSCtg", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) "
  GRP_REP.Add "Rm Code", "OrRmCtg+OrRmSCtg+OrRmCd", "'('+ RTrim(OrRmCtg)+'/'+ OrRmSCtg+ ') '+ OrRmCd", "OrRmCtg,OrRmSCtg,OrRmCd", "hOrRmCtg,hOrRmSCtg,hOrRmCd", "Rm Ctg,Rm SubCtg", "", "", "SubString(RmDesc, 1, 30) "
  GRP_REP.Add "Parent Cd", "Rtrim(RmCtg)+Rtrim(RmSCtg)+(case when RmParentCd= '' then RmCd else RmParentCd end)", "'('+ RTrim(RmCtg)+'/'+ RmSCtg+ ') '+ (case when RmParentCd= '' then RmCd else RmParentCd end)", "OrRmCtg,OrRmSCtg", "hOrRmCtg,hOrRmSCtg", "Rm Ctg,Rm SubCtg", "", "", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) "

  GRP_REP.Add "Dsg Colour", "Odm.OdDmCol", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = Odm.OdDmCol)"
  GRP_REP.Add "Karat", "Odm.OdKt", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = Odm.OdKt)"
 
 
  ElseIf UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk") Then
  GRP_REP.Add "Loc Typ", "max(LocTyp)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp = 'LOCTYP' And PMCD = max(LocTyp)) "
  GRP_REP.Add "(Typ) Loc", "RTrim(max(LocTyp))+RTrim(DCoCd)+RTrim(DLoc)", "'(' + RTrim(max(LocTyp)) + '/'+RTrim(DCoCd)+') '+ RTrim(DLoc)", "", "", "Company Code", "", "", " max(LocDesc)"
  
  GRP_REP.Add "Design Ctg", "max(DmCtg)", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= max(DmCtg)) "
  'uni.4 max(DmSalCtg) replaced by qDmSalCtgRepl,qDmSalCtgDescRepl to get either SalCtg,SalCTg2 or SalCTg3
  GRP_REP.Add "Sales Ctg", "||qDmSalCtgRepl||", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and PmCd=max(DmCtg) and PSCd= ||qDmSalCtgDescRepl|| ) "
  GRP_REP.Add "Design Code", "DDmCd", "", "", "", "Design Ctg,Sales Ctg", "", "", "max(DmDesc) "
  GRP_REP.Add "Design Size", "DDmSz", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PmCd= DDmSz)"
  GRP_REP.Add "Suffix", "DOdSfx", "", "", "", "", "", "", ""
  'sv.48 DBfldname copied to DispFldNm and DBfldname format changed
  GRP_REP.Add "DmStkNum", "DdmStkCoCd+DDmStkYy+DDmStkChr+Str(DDmStkNo)", "DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+Ltrim(Str(DDmStkNo))", "", "", "", "", "", ""
  'sv.48 In DBfldname ltrim removed from DDmStkNo
  GRP_REP.Add "DsgCd-Sfx-Sz-StkNum", "DDmCd+DOdSfx+DDmSz+DdmStkCoCd+DDmStkYy+DDmStkChr+Str(DDmStkNo)", "DDmCd+'/'+DOdSfx+'/'+DDmSz+'/'+DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+LTrim(Str(DDmStkNo))", "", "", "Design Ctg,Sales Ctg,Design Code,Design Size,Suffix,DmStkNum", "", "", ""

  If ws_MultiCoMod = True Then _
    GRP_REP.Add "Company Code", "DCoCd", "DCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=DCoCd) "

  GRP_REP.Add "Dsg Colour", "max(Odm.OdDmCol)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = max(Odm.OdDmCol))"
  GRP_REP.Add "Karat", "max(Odm.OdKt)", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = max(Odm.OdKt))"
 
 
  End If
  
End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'IdCoCd, IdTc, IdYy, IdChr, IdNo, IdSr, InExpNo, DmCtg, IdDmCd, IdSfx, IdDmSz, IdKt, InDt,
  'InCmCd, IdExpOdTc, IdExpOdYy, IdExpOdChr, IdExpOdNo, IdExpOdSr, DmSalCtg, IdVaCtg,
  'IdQty, IdIFob, IrRmIVal as qMetVal, IrRmIVal as qDCVal, IrRmIVal as qXVal,
  'IrSetIVal as qSetVal, iLiVal as qLabVal, IrRmIWt as qGrWt, IrRmIWt as qMetWt,
  'IrRmIWt As qDCWt, IrRmIWt As qXWt,
  'OdPicNm
  'from InvDsg, InvHd, DsgMst, InvRm, RmMst, InvLab, OrdDsg where 1= 2
  '*** Report Sql ***
  
  
  Dim grpflds As String, wSqlStrg As String
  Dim wrepcnd As String, wCnd As String
  
  
  Set Rep = Nothing
  Set Rep1 = Nothing
  If UCase(adc.MenuCd) = UCase("RepJtStkRm") Then
  Set adc.RepSource = Rep
  Set mRep = Rep
  ElseIf UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk") Then
  Set adc.RepSource = Rep1
  Set mRep = Rep1
  End If
  
  
  'Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(mRep)
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst + ","
  
  ws_hAddJROYN = moCn.GetFldVal("Select hAddJROYN from Head where HCoCd='ZZZ'")
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", IIF(gs_Partition = ctCurrPrtn, " and ", " ") + wrepcnd, "")
  'wCnd = wrepcnd
  ''wCnd = wCnd + " and ((DYyMM>= " + CStr(adc("UoYyMmFr")) + " and  DYyMM<= " + CStr(adc("UoYyMmTo")) + ") or qOpnQty>0) "
  
  'Uni.4 - SalCtg scope will check all of 3 Sal Ctg(SalCtg, SalCtg2 & SalCtg3).
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get where clause to join DsgMst
  Dim wDmSalCtgCnd As String, ws_DmSalCtg As String
  If adc("UoDmSalCtgFr") <> "" Or adc("UoDmSalCtgTo") <> "" Or adc("UoDmSalCtgSel") <> "" Then
    wDmSalCtgCnd = GetDmSalCtgWhrCnd(adc("UoDmSalCtgFr"), adc("UoDmSalCtgTo"), adc("UoDmSalCtgSel"), adc("wAddAllSalCtg"))
  End If
  
  'uni.4 selecting Dmsalctg either from SalCtg,SalCtg2 or SalCtg3
  'if include SalCtg2 and SalCtg3 is YEs then checks all SalCtg else behave as it is
  'calling function to get SalCtg value to show in detail section and group sort
  ws_DmSalCtg = GetDmSalCtg(adc("UoDmSalCtgFr"), adc("UoDmSalCtgTo"), adc("UoDmSalCtgSel"), adc("wAddAllSalCtg"), IIF(UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk"), "Y", ""))
  
  Dim wValDivBy As Double
  'sv.30 when divied by amount is blank then default value is 1
  wValDivBy = IIF(adc("UoValDivBy") = 0, 1, adc("UoValDivBy"))
  
  '578.22 getting base currency
  Dim ws_BaseCurCd As String, wb_BaseCur As Boolean, wCurCdCnd  As String
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + adc("UoCoCdFr") + "'")
  '578.22 if currency code not given it takes base currency from head else takes from scope
  If adc("UoCurCdFr") = "" Then
    wb_BaseCur = True
    wCurCdCnd = " "
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_BaseCur = False
    wCurCdCnd = " and OmCmCurCd= '" + adc("UoCurCdFr") + "' "
    Rep.TxtValIn.SetText "Value In " + adc("UoCurCdFr")
  End If
  
 If UCase(adc.MenuCd) = UCase("RepJtStkRm") Then
'thiru
  'sv.30 wValDivBy added
  '578.22 conversion factor added to OrCstVal
  wSqlStrg = "Select " + grpflds + " " + _
            "DCoCd, DLoc, DDmCd, DOdSfx, DDmSz, DdmStkCoCd,DDmStkYy, DDmStkChr, DDmStkNo, OrRmCd, OrLn1, OrLn2, OrRmPtr, qStkQty, " + _
            "OrRmCtg, OrRmSCtg, (OrWt * qStkQty) OrWt,(Convert(int, OrQty * qStkQty)) OrQty, " + _
            "(Convert(Decimal(14,3), OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end)* qStkQty)) as qWtInGms,   " + _
            "(Convert(Decimal(14,3), OrWt/(case when OrRmCtg in ('D', 'C') then 5 else 1 end)*qStkQty " + _
            "*(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt Else 0 end) )" + _
            "/(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then IsNull(qPurityWt, 1) Else 1 End)) as qPureWt,(qStkQty * " + IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OrSalVal/OmCnvFct else OrSalVal*OmCnvFct end) ", "OrSalVal") + ")/" + CStr(wValDivBy) + " as OrCstVal  " + _
            "From OrdRm " + _
            "           Left outer join OrdDsg OrdDsgJRO on  OrdDsgJRO.OdIdNo = OrOdIdNo and OrOdMstIdNo<>0 " + _
            "Join ( " + _
            " Select IsNull((Select OdIdNo from vwDmStkNumMaster " + _
                      "   Where MasterYN='Y' and DDmStkCoCd = OdDmStkCoCd And DDmCd = OdDmCd And DOdSfx = OdSfx And DDmSz = OdDmSz And " + _
            " DDmStkYy = OdDmStkYy and DDmStkChr = OdDmStkChr and DDmStkNo = OdDmStkNo " + IIF(gs_Partition = ctCurrPrtn, " and DPrtKey=OdPrtKey ", "") + "" + _
              "), 0) wOdIdNo , " + _
            " DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey,sum(DDmDrQty - DDmCrQty) as qStkQty " + _
            " From " + _
 "((select DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey, DDmDrQty, DDmCrQty"
 
 wSqlStrg = wSqlStrg + "  From DmYyMm where DyYMm <=" + CStr(adc("UoYyMmFr")) + ")" + _
 " Union All " + _
 "(Select OdCoCd, " + _
 "(Case when OmFrDmDc='C' or OmFrDmDc='D' then OmFrDmLoc else '' end), " + _
 " OdDmCd, OdSfx, OdDmSz, OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdPrtKey, " + _
 "(case when OmFrDmDc = 'D' then OdOrdQty else 0 end), " + _
 "(case when OmFrDmDc = 'C' then OdOrdQty else 0 end) " + _
 " From OrdDsg " + _
 " Join OrdMst on OmIdNo = OdOmIdNo " + _
 " where  OmDt >='" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "MM") + "/01", "01/" + Format(adc("UoDtFr"), "MM")) + "/" + Format(adc("UoDtFr"), "YY") + "' and OmDt<= '" + CStr(adc("UoDtFr")) + "') " + _
 " Union All " + _
 " (Select OdCoCd, " + _
 " (Case when OmToDmDc='C' or OmToDmDc='D' then OmToDmLoc else '' end), " + _
 " OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdPrtKey, " + _
 " (case when OmToDmDc = 'D' then OdOrdQty else 0 end), " + _
 " (case when OmToDmDc = 'C' then OdOrdQty else 0 end) " + _
 " From OrdDsg " + _
 " Join OrdMst on OmIdNo = OdOmIdNo" + _
 " where  OmDt >='" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "MM") + "/01", "01/" + Format(adc("UoDtFr"), "MM")) + "/" + Format(adc("UoDtFr"), "YY") + "' and OmDt<= '" + CStr(adc("UoDtFr")) + "') " + _
" ) " + _
 " " + _
 " CompleteStock "
'thiru
'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
'this is to avoid SalCtg scope condition in where clause
Dim wSqlStrg1 As String
wSqlStrg1 = wSqlStrg
wSqlStrg = wSqlStrg + "                   Group By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey Having Sum(Convert(Decimal(14,1), (DDmDrQty - DDmCrQty)))>0)DmYyMm " + _
            " On (wOdIdNo=OrOdIdNo)  " + _
            IIF(gs_Partition = ctCurrPrtn, " and DPrtkey=OrPrtKey ", "") + _
            "Join Loc On LocCoCd=DCoCd and LocCd=DLoc " + IIF(gs_Partition = ctCurrPrtn, " and LocPrtKey=DPrtKey ", "") + _
            "Join DsgMst On DmTcTyp='DM' and DmCd=DDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=DPrtKey ", "") + wDmSalCtgCnd + _
            "Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
            "join OrdDsg Odm on  Odm.OdIdNo = wOdIdNo Join OrdMst on OmIdNo = Odm.OdOmIdNo " + _
            "Left Outer Join Param SC On SC.PTyp='STNCLS' and SC.PMCd=RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
            "Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp='RMCTG')PRm " + _
            "On qPMCd= OrRmCtg " + _
            "Where /*not(DDmDrQty= 0 and DDmCrQty=0)*/ " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + _
            wCurCdCnd + " And (Select Count(*) From UsrOptDsgAna " + _
            "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
            "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
            "UdUsrCd = '" + adc.UsrCd + "' and UdMnuCd= '" + adc.MenuCd + "')= " & (GrdDsgAna.Rows - 1) & " " + _
            "/*group By DCoCd, DLoc, OrRmCd, OrLn1, OrRmPtr Having Sum(Convert(Decimal(14,1), (DDmDrQty - DDmCrQty)))>0*/  "

If ws_hAddJROYN = "Y" Then
  wSqlStrg = wSqlStrg + " union all "

  wSqlStrg = wSqlStrg + wSqlStrg1 + "                   Group By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey Having Sum(Convert(Decimal(14,1), (DDmDrQty - DDmCrQty)))>0)DmYyMm " + _
            " On (" + _
            "  OrOdMstIdNo<>0 and OrOdMstIdNo = wOdIdNo  " + _
            IIF(adc("UoDtFr") <> "01/01/80", " and isnull(OrdDsgJRO.OdOmDt,'01/01/80')<='" + CStr(adc("UoDtFr")) + "'", "") + ") " + _
            IIF(gs_Partition = ctCurrPrtn, " and DPrtkey=OrPrtKey ", "") + _
            "Join Loc On LocCoCd=DCoCd and LocCd=DLoc " + IIF(gs_Partition = ctCurrPrtn, " and LocPrtKey=DPrtKey ", "") + _
            "Join DsgMst On DmTcTyp='DM' and DmCd=DDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=DPrtKey ", "") + wDmSalCtgCnd + _
            "Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
            "join OrdDsg Odm on  Odm.OdIdNo = wOdIdNo " + _
            "Left Outer Join Param SC On SC.PTyp='STNCLS' and SC.PMCd=RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
            "Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp='RMCTG')PRm " + _
            "On qPMCd= OrRmCtg " + _
            "Where OrOdMstIdNo<>OrOdIdNo and /*not(DDmDrQty= 0 and DDmCrQty=0)*/ " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + _
             " And (Select Count(*) From UsrOptDsgAna " + _
            "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
            "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
            "UdUsrCd = '" + adc.UsrCd + "' and UdMnuCd= '" + adc.MenuCd + "')= " & (GrdDsgAna.Rows - 1) & " " + _
            "/*group By DCoCd, DLoc, OrRmCd, OrLn1, OrRmPtr Having Sum(Convert(Decimal(14,1), (DDmDrQty - DDmCrQty)))>0*/  Order By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz, DDmStkYy, DDmStkChr, DDmStkNo"

End If
ElseIf UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk") Then

'thiru
'sv.8 Pure wt calculation changed, OrWt * qStkQty added to pure wt
   wSqlStrg = "Select " + grpflds + " " + _
            "DCoCd,max(LocTyp) as LocTyp, DLoc, DDmCd, DOdSfx, DDmSz, DdmStkCoCd,DDmStkYy, DDmStkChr, DDmStkNo,max(Odm.OdDmCol) as OdDmCol, max(Odm.OdKt) as OdKt,max(qStkQty) as qStkQty,Datediff(Day,max(odm.OdOmDt), getdate() ) as qStkDay, " + _
            "max(odm.OdSalprc) as qStkVal,  " + _
            "sum(case when    IsNull(OrRmCtg, '') in ('D','C') then  IsNull( OrWt, 0)/5 else  IsNull( OrWt, 0) end) as qGrsWt ,sum((case when    IsNull(OrRmCtg, '') in ('G','P','S','L','M')  then  IsNull( OrWt, 0) else 0 end)* qStkQty) as qMetWt , " + _
            "sum(Convert(Decimal(14,3),IsNull(OrWt, 0) * qStkQty*(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then RmPurityWt Else 0 end) )" + _
            "/(case when OrRmCtg in ('G', 'P', 'S', 'L', 'M') then IsNull(qPurityWt, 1) Else 1 End)) as qPureWt, " + _
            "sum((case when    IsNull(OrRmCtg, '') ='D' then  IsNull( OrWt, 0) else 0 end) * qStkQty )  as qDiaWt," + _
            "sum((case when    IsNull(OrRmCtg, '') ='D' then OrQty else 0 end)* qStkQty)  as qDiaQty ,sum((case when    IsNull(OrRmCtg, '') ='C' then IsNull( OrWt, 0) else 0 end)* qStkQty)  as qCsWt,sum((case when    IsNull(OrRmCtg, '') ='C' then OrQty else 0 end)* qStkQty)  as qCsQty,sum((case when    IsNull(OrRmCtg, '') ='X' then  IsNull( OrWt, 0) else 0 end)* qStkQty) as qXWt " + _
            "From OrdRm " + _
            "Left outer join OrdDsg OrdDsgJRO on  OrdDsgJRO.OdIdNo = OrOdMstIdNo and OrOdMstIdNo<>0 " + _
            "Join ( " + _
            " Select IsNull((Select (OdIdNo) from vwDmStkNumMaster " + _
            "Where MasterYN='Y' and DDmCd = OdDmCd And DOdSfx = OdSfx And DDmSz = OdDmSz And DdmStkCoCd=OdDmStkCoCd and DDmStkYy = OdDmStkYy and DDmStkChr = OdDmStkChr and  DDmStkNo = OdDmStkNo ), 0) wOdIdNo , " + _
            " DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey,sum(DDmDrQty - DDmCrQty) as qStkQty " + _
            " From " + _
 "((select DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey, DDmDrQty, DDmCrQty "
 
 
 wSqlStrg = wSqlStrg + "  From DmYyMm where DyYMm <=" + CStr(adc("UoYyMmFr")) + ")" + _
 " Union All " + _
 "(Select OdCoCd, " + _
 "(Case when OmFrDmDc='C' or OmFrDmDc='D' then OmFrDmLoc else '' end), " + _
 " OdDmCd, OdSfx, OdDmSz, OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdPrtKey, " + _
 "(case when OmFrDmDc = 'D' then OdOrdQty else 0 end),(case when OmFrDmDc = 'C' then OdOrdQty else 0 end) " + _
 " From OrdDsg " + _
 " Join OrdMst on OmIdNo = OdOmIdNo " + _
 " where  OmDt >='" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "MM") + "/01", "01/" + Format(adc("UoDtFr"), "MM")) + "/" + Format(adc("UoDtFr"), "YY") + "' and OmDt<= '" + CStr(adc("UoDtFr")) + "') " + _
 " Union All " + _
 " (Select OdCoCd, " + _
 " (Case when OmToDmDc='C' or OmToDmDc='D' then OmToDmLoc else '' end), " + _
 " OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdPrtKey, " + _
 " (case when OmToDmDc = 'D' then OdOrdQty else 0 end), (case when OmToDmDc = 'C' then OdOrdQty else 0 end)" + _
 " From OrdDsg " + _
 " Join OrdMst on OmIdNo = OdOmIdNo  " + _
 " where  OmDt >='" + IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(adc("UoDtFr"), "MM") + "/01", "01/" + Format(adc("UoDtFr"), "MM")) + "/" + Format(adc("UoDtFr"), "YY") + "' and OmDt<= '" + CStr(adc("UoDtFr")) + "') " + _
" ) " + _
 " " + _
 " CompleteStock "
'            IIF(ADC("UoDtFr") <> "01/01/80", " and (Select OdOmDt from OrdDsg where OdIdNo=OrOdIdNo)<='" + CStr(ADC("UoDtFr")) + "'", "") + ")) " + _

'thiru
'uni.4 wDmSalCtgCnd added to DsgMst join, to get SalCtg value from scope.
'this is to avoid SalCtg scope condition in where clause
wSqlStrg = wSqlStrg + "                   Group By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DPrtKey Having Sum(Convert(Decimal(14,1), (DDmDrQty - DDmCrQty)))>0)DmYyMm " + _
            " On ((wOdIdNo=OrOdIdNo) " + _
            " or ('Y' = '" + ws_hAddJROYN + "' and OrOdMstIdNo <>0 and OrOdMstIdNo = wOdIdNo " + _
            IIF(adc("UoDtFr") <> "01/01/80", " and isnull(OrdDsgJRO.OdOmDt,'01/01/80')<='" + CStr(adc("UoDtFr")) + "'", "") + ")) " + _
            IIF(gs_Partition = ctCurrPrtn, " and DPrtkey=OrPrtKey ", "") + _
            "Join Loc On LocCoCd=DCoCd and LocCd=DLoc " + IIF(gs_Partition = ctCurrPrtn, " and LocPrtKey=DPrtKey ", "") + _
            "Join DsgMst On DmTcTyp='DM' and DmCd=DDmCd and DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=DPrtKey ", "") + wDmSalCtgCnd + _
            "Join RmMst On RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
            " join OrdDsg Odm on Odm.OdIdNo = wOdIdNo " + _
            "Left Outer Join Param SC On SC.PTyp='STNCLS' and SC.PMCd=RmCtg and ',' + SC.PDesc225 + ',' like '%,'+ RmSCtg +',%' " + _
            "Left Outer Join (Select PMCd as qPMCd, RmPurityWt as qPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp='RMCTG')PRm " + _
            "On qPMCd= OrRmCtg " + _
            "Where " + IIF(gs_Partition = ctCurrPrtn, " DPrtKey='" + ctCurrPrtn + "' ", "") + wCnd + _
             " And (Select Count(*) From UsrOptDsgAna " + _
            "Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
            "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) and DaTcTyp = DmTcTyp and DaCd= DmCd and DaSz= DmSz and " + _
            "UdUsrCd = '" + adc.UsrCd + "' and UdMnuCd= '" + adc.MenuCd + "')= " & (GrdDsgAna.Rows - 1) & " " + _
            "group By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo  Order By DCoCd, DLoc, DDmCd, DOdSfx, DDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo"
 End If
  
 'uni.4 replacing DmSalCtg value and its description in group sort of Sales Ctg
 wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgRepl||", ws_DmSalCtg)
 wSqlStrg = Replace(wSqlStrg, "||qDmSalCtgDescRepl||", ws_DmSalCtg)
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt
  
  ShowRepSummFld Rep.Fldr1StkQty, Rep.Fldr2StkQty, Rep.Fldr3StkQty, Rep.FldrStkQty, GRP_REP.GrpOpts, "Rm Ctg,Rm SubCtg,Rm Code,Parent Cd"
  
  If ws_MultiCoMod Then
    mRep.TxtCoCdFr.SetText adc("UoCoCdFr"): mRep.TxtCoCdTo.SetText adc("UoCoCdTo"): mRep.TxtCoCdSel.SetText adc("UoCoCdSel")
  Else
    mRep.TxtCoCdToLbl.Suppress = True: mRep.TxtCoCdSelLbl.Suppress = True
    mRep.TxtCoCdFr.SetText adc("UoCoCdFr")
  End If
  mRep.TxtYyMmFr.SetText adc("UoYyMmFr")
  mRep.txtLocTypFr.SetText adc("UoLocTypFr")
  mRep.TxtLocTypTo.SetText adc("UoLocTypTo")
  mRep.TxtLocTypSel.SetText adc("UoLocTypSel")
  
  If UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk") Then
    If adc("UoYN") = "Y" Then
    Rep1.TxtGrsWt.Suppress = True: Rep1.TxtGrsWt1.Suppress = True: Rep1.qGrsWt1.Suppress = True: Rep1.gGrsWt1.Suppress = True
    Rep1.g1GrsWt1.Suppress = True: Rep1.g2GrsWt1.Suppress = True: Rep1.g3GrsWt1.Suppress = True

    Rep1.TxtMetWt.Suppress = True: Rep1.TxtMetWt1.Suppress = True: Rep1.qMetWt1.Suppress = True: Rep1.gMetWt1.Suppress = True
    Rep1.g1MetWt1.Suppress = True: Rep1.g2MetWt1.Suppress = True: Rep1.g3MetWt1.Suppress = True

    Rep1.TxtPureWt.Suppress = True: Rep1.TxtPureWt1.Suppress = True: Rep1.qPureWt1.Suppress = True: Rep1.gPureWt1.Suppress = True
    Rep1.g1PureWt1.Suppress = True: Rep1.g2PureWt1.Suppress = True: Rep1.g3PureWt1.Suppress = True

    Rep1.TxtDiaWt.Suppress = True: Rep1.TxtDiaWt1.Suppress = True: Rep1.qDiaWt1.Suppress = True: Rep1.gDiaWt1.Suppress = True
    Rep1.g1DiaWt1.Suppress = True: Rep1.g2DiaWt1.Suppress = True: Rep1.g3DiaWt1.Suppress = True

    Rep1.TxtCsWt.Suppress = True: Rep1.TxtCsWt1.Suppress = True: Rep1.qCsWt1.Suppress = True: Rep1.gCsWt1.Suppress = True
    Rep1.g1CsWt1.Suppress = True: Rep1.g2CsWt1.Suppress = True: Rep1.g3CsWt1.Suppress = True

    Rep1.TxtXWt.Suppress = True: Rep1.TxtXWt1.Suppress = True: Rep1.qXWt1.Suppress = True: Rep1.gXWt1.Suppress = True
    Rep1.g1xwt1.Suppress = True: Rep1.g2XWt1.Suppress = True: Rep1.g3XWt1.Suppress = True
    
    Rep1.TxtHead.SetText "Design Flash Stock (Jwlry) - Value"
    
    
    ElseIf adc("UoYN") = "N" Then
    
    Rep1.TxtPerPcVal.Suppress = True: Rep1.TxtPerPcVal1.Suppress = True:   Rep1.qStkVal1.Suppress = True: Rep1.gStkVal1.Suppress = True
    Rep1.g1StkVal1.Suppress = True: Rep1.g2StkVal1.Suppress = True: Rep1.g3Stkval1.Suppress = True
    Rep1.TxtHead.SetText "Design Flash Stock (Jwlry) - Weight"
    End If
    
  End If
  'sv.30 passing divide by value to dsr
  GRP_REP.SetFormula Rep, "wDivByVal", adc("UoValDivBy")
  
  'sv.30 when divied by amount>1 then follwing text will display in header
  If adc("UoValDivBy") > 1 Then Rep.TxtDivByVal.SetText "(Values in " + Format(adc("UoValDivBy"), "###,##0") + ")"
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.OrRmCd}", 7)
  
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc)

End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  'uni.4  checking Include SalCtg2,SalCtg3 option
  Case Is = 0
    If .Value = Checked Then adc("wAddAllSalCtg") = "Y" Else adc("wAddAllSalCtg") = "N"
  End Select
End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
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
  
  'PG.7 By Default report show as weight only
  OptValWt(1).Value = True
  
  '*** Jenny Colour
End Sub
Private Sub ADC_Load()
  Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
  
  'pg
  adc.Child = "GrdDsgAna"
  Set GrdDsgAna.PrevCtl = adc
  Set GrdDsgAna.NextCtl = adc
  GrdDsgAna.HotKey = "A"
  
  
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  
  If ws_MultiCoMod = True Then
    Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Else
    Set adc.FirNKeyCtl = adc("UOLOCTYPFR")
  End If

  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  If ws_MultiCoMod = True Then
      adc("UOCOCDFR").Enabled = True: adc("UOCOCDTO").Visible = True: adc("UOCOCDSEL").Visible = True
      adc("UOCOCDTO").Enabled = True: adc("UOCOCDSEL").Enabled = True
      adc("UoPrntCdFr").Enabled = True: adc("UoPrntCdFr").Visible = True
      adc("UoPrntCdTo").Enabled = True: adc("UoPrntCdTo").Visible = True
      adc("UoPrntCdSel").Enabled = True: adc("UoPrntCdSel").Visible = True
  Else
      adc("UOCOCDFR").Enabled = False: adc("UOCOCDTO").Visible = False: adc("UOCOCDSEL").Visible = False
      adc("UoPrntCdFr").Enabled = False: adc("UoPrntCdFr").Visible = False
      adc("UoPrntCdTo").Enabled = False: adc("UoPrntCdTo").Visible = False
      adc("UoPrntCdSel").Enabled = False: adc("UoPrntCdSel").Visible = False
  End If
  
   If UCase(adc.MenuCd) = UCase("RepJtStkRm") Then
   LblVal.Visible = False
   FraOptValWt.Enabled = False: FraOptValWt.Visible = False
   adc("UoYN").Enabled = False: adc("UoYN").Visible = False
   End If

'PG.7 Disabled RmCtg , Stone Class, Rm Sub Ctg, Rm Code, Rm Ln/ Sv, Breadth
   If UCase(adc.MenuCd) = UCase("RepJtDsgFlashStk") Then
   
    LblRmCtg.Visible = False
    adc("UoRmCtgFr").Visible = False: adc("UoRmCtgTo").Visible = False: adc("UoRmCtgSel").Visible = False
    adc("UoRmCtgFr").Enabled = False: adc("UoRmCtgTo").Enabled = False: adc("UoRmCtgSel").Enabled = False
    LblStnCls.Visible = False
    adc("UoStnClsFr").Visible = False: adc("UoStnClsTo").Visible = False: adc("UoStnClsSel").Visible = False
    adc("UoStnClsFr").Enabled = False: adc("UoStnClsTo").Enabled = False: adc("UoStnClsSel").Enabled = False
    LblRmSubCtg.Visible = False
    adc("UoRmSCtgFr").Visible = False: adc("UoRmSCtgTo").Visible = False: adc("UoRmSCtgSel").Visible = False
    adc("UoRmSCtgFr").Enabled = False: adc("UoRmSCtgTo").Enabled = False: adc("UoRmSCtgSel").Enabled = False
    LblRmCd.Visible = False
    adc("UoRmCdFr").Visible = False: adc("UoRmCdTo").Visible = False: adc("UoRmCdSel").Visible = False
    adc("UoRmCdFr").Enabled = False: adc("UoRmCdTo").Enabled = False: adc("UoRmCdSel").Enabled = False
    LblLen.Visible = False
    adc("UoLn1Fr").Visible = False: adc("UoLn1To").Visible = False
    adc("UoLn1Fr").Enabled = False: adc("UoLn1To").Enabled = False
    LblBrdth.Visible = False
    adc("UoLn2Fr").Visible = False: adc("UoLn2To").Visible = False
    adc("UoLn2Fr").Enabled = False: adc("UoLn2To").Enabled = False
    
   End If
  
  Call SetGroupSort
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gb_CoCdFor = Not ws_MultiCoMod
  'sv.30 divided by amount cannot be 0
  Select Case UCase(IdName)
  Case Is = UCase("UoValDivBy")
    If pv_NewValue = 0 Then Cancel = True: ErrMsg = "Divided By Amount Cannot be Zero": Exit Sub
  End Select
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  If IdName = "UODTFR" Then
    If pv_NewValue <> "01/01/80" Then
      adc("UOYYMMFR") = Format(DateAdd("m", -1, CDate(pv_NewValue)), "YYYYMM")
    End If
'  wFrDt = "01/" + Format(adc("UoDtFr"), "mm") + "/" + Format(adc("UoDtFr"), "yy")
  'If moCn.RecSeek("Select 'x' where convert(varchar(15),DateAdd(Day,-1,DateAdd(Month ,1,'01/'+Cast(Month('" + CStr(wDtTo) + "') as Varchar(2))+'/'+Cast(Year('" + CStr(wDtTo) + "')as Varchar(4)))),3)='" + wDtTo + "'") Then wb_MonEnd = True
  End If
    
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gb_CoCdFor = Not ws_MultiCoMod
    
  ' PG.7 Location and parent location help for corresponding location typ
  Select Case UCase(IdName)
  Case Is = UCase$("UoIssLocFr")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
  Case Is = UCase$("UoIssLocTo")
    gs_LocTyp = IIF(adc("UoLocTypTo") = "", "", "'" + adc("UoLocTypTo") + "'")
  Case Is = UCase$("UoIssLocSel")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
  Case Is = UCase$("UoPrntCdFr")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
  Case Is = UCase$("UoPrntCdTo")
    gs_LocTyp = IIF(adc("UoLocTypTo") = "", "", "'" + adc("UoLocTypTo") + "'")
  Case Is = UCase$("UoPrntCdSel")
    gs_LocTyp = IIF(adc("UoLocTypFr") = "", "", "'" + adc("UoLocTypFr") + "'")
  
  End Select
  
  SetProp Me, IdName, When
  If IdName = "UOYYMMFR" Then
    Cancel = False
    If adc("UODTFR") <> "01/01/80" Then
      Cancel = True
      ErrMsg = "Make As Of Dt as 01/01/80 to edit YyMm"
    End If
  End If
  
   
  'gs_LocTyp = "'F'"
 
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

  If ws_MultiCoMod = False Then _
      adc("UoCoCdFr") = gs_CoCd: adc("UoCoCdTo") = gs_CoCd: adc("UoCoCdSel") = "": adc("UoCoCdFr").Enabled = False
  
  adc("UoDtFr") = "01/01/80"
  
'Thirumozhi
  If adc("UoYN") = "Y" Then
  OptValWt(0).Value = True
  End If
  
  If adc("UoYN") = "N" Then
  OptValWt(1).Value = True
  End If
  'uni.4 default value for Include SalCtg2,SalCtg3 is Yes
  ChkBoxArr(0).Value = Checked
  adc("wAddAllSalCtg") = "Y"
  adc("UoValDivBy") = 1     'sv.30 setting default value
  'sv.30 divide by value will be visible only for jemr Stock Rm report
  If UCase(adc.MenuCd) <> UCase("RepJtStkRm") Then
    LblValDivBy.Visible = False
    adc("UoValDivBy").Visible = False
    LblCurCd.Visible = False: adc("UoCurCdFr").Visible = False  '578.22 currency code will be visible only for jemr Stock Rm report
  End If

End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, ctSelfCoCd, adc("UOCOCDFR")) + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  'Rep.wCoCd.SetText ws_HName
  mRep.wCoCd.SetText ws_HName
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
 ' Rep.wCoCdLogo.SetText ws_HName
  mRep.wCoCdLogo.SetText ws_HName
  
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    mRep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    mRep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
  
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  mRep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

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
'pg.6- Set visible /invisible state for analysis frame
Private Sub CmdSc_Click()
  If FraSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSc.Enabled = True
    FraSc.Visible = True
    TAB_REP.Enabled = False
    adc.AllowSave = False
    FraSc.ZOrder
  Else
    TAB_REP.Enabled = True
    adc.AllowSave = True
    FraSc.Enabled = False
    FraSc.Visible = False
  End If
End Sub

'pg.6-checking Valid Analysis Sr & field
Private Sub GrdDsgAna_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaSr")
      Cancel = Not moCn.RecSeek("Select PMCd From Param Where " + _
                                "PTyp= 'DAANAFLD' And PMCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Sr.": Exit Sub
    Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
      Cancel = NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param Where PTyp= 'DAANACD' And PMCd= '" + _
                                GrdDsgAna.Value(RowNum, "UdAnaSr") + "'And PSCd = '" + NewValue + "'")
      ErrMsg = "Invalid Analysis Field ": Exit Sub
    End Select
  End With
End Sub
'pg- calling hlp for GrdDsgAna
Private Sub GrdDsgAna_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  Select Case ColName
  Case Is = UCase("UdAnaSr")
    Call HlpList.PMCd("DAANAFLD")
  Case Is = UCase("UdAnaCdFr"), UCase("UdAnaCdTo")
    If moCn.GetFldVal("Select PValue from Param where PTyp= 'DAANAFLD' and " + _
       "PMCd= '" + GrdDsgAna.Value(RowNum, "UdAnaSr") + "' ") = "N" Then _
       Cancel = True: ErrMsg = "Cannot Enter Code For This Analysis Field": Exit Sub
  
    Call HlpList.PSCd("DAANACD", GrdDsgAna.Value(RowNum, "UdAnaSr"))
  End Select
End Sub

Private Sub GrdDsgAna_InitKey(ByVal RowNum As Integer)
  With GrdDsgAna
    If .IsNew(RowNum) Then
      .Store "UdUsrCd", adc.UsrCd
      .Store "UdMnuCd", adc.MenuCd
    End If
  End With
End Sub

Private Sub GrdDsgAna_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  With GrdDsgAna
    Select Case UCase(ColName)
    Case Is = UCase("UdAnaCdFr")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdFr") = ""
    Case Is = UCase("UdAnaCdTo")
      If .Mode = fgmnorm Then .Value(RowNum, "UdAnaCdTo") = ""
    End Select
  End With
End Sub

Private Sub GrdDsgAna_SetRecSource()
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + adc.UsrCd + "' and " + _
                        "UdMnuCd= '" + adc.MenuCd + "'"
                        
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
  Cancel = GrdDsgAna.Validate
End Sub

''Thirumozhi
'Private Sub OptValWt_Click(Index As Integer)
'Select Case Index
'Case Is = 0
'  adc("UoYN") = "Y"
'Case Is = 1
'  adc("UoYN") = "N"
'
'End Select
'End Sub


