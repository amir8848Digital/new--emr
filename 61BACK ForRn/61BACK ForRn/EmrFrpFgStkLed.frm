VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "MwfCtlForSpeed610.ocx"
Begin VB.Form EmrFrpFgStkLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Ledger"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdSc 
      Height          =   390
      Left            =   2160
      TabIndex        =   66
      ToolTipText     =   "Analysis Scope Options"
      Top             =   9700
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
      Left            =   11010
      TabIndex        =   38
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
      TabIndex        =   39
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
      Left            =   4965
      TabIndex        =   48
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   49
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
         TabIndex        =   50
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
      Left            =   6180
      TabIndex        =   47
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   44
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
      TabIndex        =   40
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
      TabIndex        =   43
      Top             =   60
      Width           =   15255
      Begin VB.Frame FraSc 
         Height          =   2925
         Left            =   1320
         TabIndex        =   67
         Top             =   6600
         Visible         =   0   'False
         Width           =   3945
         Begin MwfCtl.MWCTL_FLX GrdDsgAna 
            Height          =   2430
            Left            =   120
            TabIndex        =   68
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
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   42
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpFgStkLed.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpFgStkLed.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   46
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   41
               Top             =   360
               Width           =   10245
               _ExtentX        =   18071
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   45
            Top             =   360
            Width           =   15045
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
               Index           =   1
               Left            =   30
               TabIndex        =   34
               Tag             =   "ShowPic"
               ToolTipText     =   "Check To Show Picture"
               Top             =   4275
               Width           =   1900
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Reverse Logic    "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   30
               TabIndex        =   33
               Tag             =   "ShowLine"
               ToolTipText     =   "Select Reverse Logic"
               Top             =   3960
               Width           =   1900
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4275
               TabIndex        =   10
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   1560
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
               CmpStr          =   "LocCd<= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   6690
               TabIndex        =   11
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   1560
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
               CmpStr          =   "LocCd In "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4275
               TabIndex        =   31
               ToolTipText     =   "Enter To Voucher Date"
               Top             =   3345
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
               Left            =   1740
               TabIndex        =   30
               ToolTipText     =   "Enter From Voucher Date"
               Top             =   3345
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
               Left            =   1740
               TabIndex        =   9
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   1560
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
               CmpStr          =   "LocCd>= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1740
               TabIndex        =   32
               ToolTipText     =   "Enter Open YearMonth"
               Top             =   3630
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
               Index           =   24
               Left            =   2280
               TabIndex        =   37
               ToolTipText     =   "Enter 'Y' to apply Reverse Logic else enter 'N' for normal logic"
               Top             =   3945
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
               Index           =   47
               Left            =   4275
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "LocCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   6690
               TabIndex        =   2
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   705
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "LocCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   4275
               TabIndex        =   7
               ToolTipText     =   "Enter To Parent Code"
               Top             =   1275
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
               Index           =   49
               Left            =   6690
               TabIndex        =   8
               ToolTipText     =   "Enter Parent Code Selection"
               Top             =   1275
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRNTCDSEL"
               IdName          =   "UOPRNTCDSEL"
               CmpStr          =   "LocPrntCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1740
               TabIndex        =   6
               ToolTipText     =   "Enter From Parent Code"
               Top             =   1275
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
               Index           =   4
               Left            =   1740
               TabIndex        =   15
               ToolTipText     =   "Enter From Design Code"
               Top             =   2145
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4275
               TabIndex        =   16
               ToolTipText     =   "Enter To Design Code"
               Top             =   2145
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1740
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Sufix"
               Top             =   2430
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXFR"
               IdName          =   "UOODSFXFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   7
               Left            =   4275
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Sufix"
               Top             =   2430
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UOODSFXTO"
               IdName          =   "UOODSFXTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   6690
               TabIndex        =   17
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   2145
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   3360
               TabIndex        =   25
               ToolTipText     =   "Enter From Stock Number"
               Top             =   3045
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDFR"
               IdName          =   "UOODDMIDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   5940
               TabIndex        =   29
               ToolTipText     =   "Enter To Stock Number"
               Top             =   3045
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Mask            =   "#0"
               MaxLength       =   2
               DataType        =   1
               DataField       =   "UOODDMIDTO"
               IdName          =   "UOODDMIDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   4275
               TabIndex        =   21
               ToolTipText     =   "Enter To Design Size"
               Top             =   2715
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZTO"
               IdName          =   "UODMSZTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1740
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   705
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "LocCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1740
               TabIndex        =   3
               ToolTipText     =   "Enter From Location Type"
               Top             =   990
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
               Left            =   4275
               TabIndex        =   4
               ToolTipText     =   "Enter To Location Type"
               Top             =   990
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
               Index           =   28
               Left            =   6690
               TabIndex        =   5
               ToolTipText     =   "Enter Location Type Selection"
               Top             =   990
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOLOCTYPSEL"
               IdName          =   "UOLOCTYPSEL"
               CmpStr          =   "LocTyp  in "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   4860
               TabIndex        =   27
               ToolTipText     =   "Enter To Stock Year"
               Top             =   3045
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYTO"
               IdName          =   "UOBYYTO"
               CmpStr          =   "OdYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2640
               TabIndex        =   24
               ToolTipText     =   "Enter From Stock Character"
               Top             =   3045
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRFR"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "OdChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   5205
               TabIndex        =   28
               ToolTipText     =   "Enter To Stock Character"
               Top             =   3045
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOBCHRTO"
               IdName          =   "UOBCHRTO"
               CmpStr          =   "OdChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1740
               TabIndex        =   20
               ToolTipText     =   "Enter From Design Size"
               Top             =   2715
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UODMSZFR"
               IdName          =   "UODMSZFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2340
               TabIndex        =   23
               ToolTipText     =   "Enter From Stock Year"
               Top             =   3045
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOBYYFR"
               IdName          =   "UOBYYFR"
               CmpStr          =   "OdYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1740
               TabIndex        =   22
               ToolTipText     =   "Enter From Stock company code"
               Top             =   3045
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDFR"
               IdName          =   "UOSTKCOCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   4275
               TabIndex        =   26
               ToolTipText     =   "Enter To Stock company code"
               Top             =   3045
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOSTKCOCDTO"
               IdName          =   "UOSTKCOCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2280
               TabIndex        =   64
               ToolTipText     =   "Show Picture? ( Enter 'Y'/'N')"
               Top             =   4275
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
               Index           =   15
               Left            =   4275
               TabIndex        =   13
               ToolTipText     =   "Enter To Design Category"
               Top             =   1860
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
               CmpStr          =   "DmCtg<="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   1740
               TabIndex        =   12
               ToolTipText     =   "Enter From Design Category"
               Top             =   1860
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
               CmpStr          =   "DmCtg>="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   76
               Left            =   6690
               TabIndex        =   14
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   1860
               Width           =   8280
               _ExtentX        =   14605
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
               CmpStr          =   "DmCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1740
               TabIndex        =   35
               ToolTipText     =   "Enter Divided By Amount "
               Top             =   4680
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
               Index           =   44
               Left            =   1740
               TabIndex        =   36
               ToolTipText     =   "Enter Currency Code Whose Orders You Wish To View Or Leave It Blank If You Wish To View All Orders In Base Currency"
               Top             =   4965
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
               TabIndex        =   70
               Top             =   4965
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
               TabIndex        =   69
               Top             =   4680
               Width           =   1665
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
               TabIndex        =   65
               Top             =   1860
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
               Left            =   60
               TabIndex        =   63
               Top             =   2715
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
               Index           =   0
               Left            =   60
               TabIndex        =   62
               Top             =   990
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
               Index           =   30
               Left            =   60
               TabIndex        =   61
               Top             =   2145
               Width           =   1425
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
               Index           =   31
               Left            =   60
               TabIndex        =   60
               Top             =   2430
               Width           =   1425
            End
            Begin VB.Label ALBL 
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
               Index           =   32
               Left            =   60
               TabIndex        =   59
               Top             =   3045
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
               Left            =   60
               TabIndex        =   58
               Top             =   1275
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
               TabIndex        =   57
               Top             =   3615
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
               TabIndex        =   56
               Top             =   705
               Width           =   1335
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
               TabIndex        =   55
               Top             =   3330
               Width           =   1425
            End
            Begin VB.Label LblLocCd 
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
               TabIndex        =   54
               Top             =   1560
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
               Left            =   6660
               TabIndex        =   53
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
               Left            =   4275
               TabIndex        =   52
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
               Left            =   1740
               TabIndex        =   51
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpFgStkLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepFgStkLed
Dim moCn As MwfLib.MDOConnection
Dim ws_MultiCoMod As Boolean
Private Sub SetGroupSort()
  GRP_REP.Add "Loc Typ", "LocTyp", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp = 'LOCTYP' And PMCD = LocTyp) "
  GRP_REP.Add "(Typ) Loc", "RTrim(LocTyp)+RTrim(LocCoCd)+RTrim(LocCd)", "'(' + RTrim(LocTyp) + '/'+RTrim(LocCoCd)+') '+ RTrim(LocCd)", "qLoc", "hqLoc", "Company Code,Loc Typ", "", "", " LocDesc"

  GRP_REP.Add "Location", "LocCoCd+LocCd", "'['+LocCoCd+'] '+LocCd", "qLoc", "hqLoc", "Company Code,Parent Code", "", "", "LocDesc"
  Dim wsPtrYN As String
  If ws_MultiCoMod = True Then
    GRP_REP.Add "Company Code", "LocCoCd", "LocCoCd", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=LocCoCd) "
    GRP_REP.Add "Parent Location", "LocPrntCd", "LocPrntCd", "", "", "", "", "", "(Select LocDesc From Loc Where LocCoCd= '" + ctSelfCoCd + "' And LocCd=Lc.LocPrntCd) "
  End If
  'GRP_REP.Add "Suffix", Array("DDmCd+DOdSfx", "OdDmCd+OdSfx"), Array("'('+DDmCd+')'+DOdSfx", "'('+OdDmCd+')'+OdSfx"), "OdDmCd,OdSfx", "hDmCd,hOdSfx", "Design Code", "", "", Array("''", "OdSfxDesc ")
  GRP_REP.Add "Suffix", Array("DOdSfx", "OdSfx"), Array("DOdSfx", "OdSfx"), "OdSfx", "hOdSfx", "", "", "", Array("''", "OdSfxDesc ")
  GRP_REP.Add "Design Code", Array("DDmCd", "OdDmCd"), Array("DDmCd", "OdDmCd"), "OdDmCd", "HDmCd", "", "", "", ""
  'sv.48 In DBfldname ltrim removed from DDmStkNo & OdDmStkNo
  GRP_REP.Add "DsgStkNum", Array("DdmStkCoCd+DDmStkYy+DDmStkChr+str(DDmStkNo)", "OdDmStkCoCd+OdDmStkYy+OdDmStkChr+str(OdDmStkNo)"), Array("DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+LTrim(str(DDmStkNo))", "OdDmStkCoCd+'/'+OdDmStkYy+'/'+OdDmStkChr+'/'+LTrim(str(OdDmStkNo))"), "", "", "", "", "", ""
  'sv.48 In DBfldname ltrim removed from DDmStkNo & OdDmStkNo
  GRP_REP.Add "DsgCd-StkNo", Array("DDmCd+DdmStkCoCd+DDmStkYy+DDmStkChr+str(DDmStkNo)", "OdDmCd+OdDmStkCoCd+OdDmStkYy+OdDmStkChr+str(OdDmStkNo)"), Array("'('+DDmCd+')'+DdmStkCoCd+'/'+DDmStkYy+'/'+DDmStkChr+'/'+LTrim(str(DDmStkNo))", "'('+OdDmCd+')'+OdDmStkCoCd+'/'+OdDmStkYy+'/'+OdDmStkChr+'/'+LTrim(str(OdDmStkNo))"), "", "", "Design Code,Suffix,StkNo", "", "", ""
  
  'PG.8- DsgCtg group sort added
  GRP_REP.Add "Design Category", "DmCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp = 'DMCTG' And PMCD = DmCtg) "
    
End Sub
Private Sub adc_setreprecsource()
On Error GoTo errhld
  '*** Report Sql ***
  
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  space(1) as qFlag, OdDmCd, OdSfx, OdDmSz, OdDmId, OdCoCd, OdTc, OdYy,
  '  OdChr, OdNo, OdSr, OmDt as qDt, OmPoNo as qDesc, LocPrntCd as qPrntCd,
  '  LocCd as qLoc, (DDmDrQty - DDmCrQty) as qOpnQty, DDmDrQty as qDrQty,
  '  DDmCrQty As qCrQty, OmFrDmLoc As qFrDmLoc, OmToDmLoc As qToDmLoc
  '  Into TmpFgStkLed From DmYyMm, OrdMst, OrdDsg, Loc, Param where 1=2
  
  Dim grpflds0 As String, grpflds1 As String, grpflds2 As String, grpflds3 As String, wrepcnd As String
  Dim wCnd As String, wDtCndOpnYyMm As String, wOpnYyStr As String
  Dim wsqlOpnYyMm As String, wFrDt As String, wDtCndOpnOd As String, wSqlOpnTdFr As String
  Dim wSqlOpnTdTo As String, wDtCndDet As String, wSqlDetFr As String, wSqlDetTo, wSqlFinal As String
  Dim wTmpDt As String, wSetTmpTab As String, wRLQty As String
  Dim ws_DsgAnaCnd As String, wCurCdCnd As String

  Dim ws_PurgeDt As String
  'sv.9 getting PurgeDate from vParam ,for jEmr taking from ZZZ company for others taking from given company code
  If UCase(ADC.MenuCd) = UCase("RepJtStkLed") Then
    ws_PurgeDt = moCn.GetFldVal("Select vPDesc225 from vParam where vPCoCd ='ZZZ' and " + _
                                " vPTyp= 'JPURGEDT' and vPMCd= 'JPURGEDT' and vPSCd =''")
  
  Else
    ws_PurgeDt = moCn.GetFldVal("Select max(convert(date,vPDesc225,3)) from vParam " + _
                                " where vPCoCd >='" + ADC("UoCoCdFr") + "' and " + _
                                " vPCoCd <='" + IIF(ADC("UoCoCdTo") <> "", ADC("UoCoCdTo"), "ZZZ") + "'" + _
                                IIF(ADC("UoCoCdSel") <> "", " and vPCoCd In(" + ADC("UoCoCdSel") + ")", "") + _
                                " and vPTyp= 'PURGEDT' and vPMCd= 'PURGEDT' and vPSCd =''")
  End If
  
  If ws_PurgeDt = "" Then ws_PurgeDt = "01/01/80"
  
  'sv.9 gs_PurgeDt replaced with ws_PurgeDt
  If gs_Partition = ctCurrPrtn And CDate(ws_PurgeDt) <> "01/01/80" And ((CDate(ADC("UOOPNYYMM")) <= Format(CDate(ws_PurgeDt), "YYYYMM") And ADC("UOOPNYYMM") <> 0) Or CDate(ADC("UODTFR")) <= CDate(ws_PurgeDt) Or CDate(ADC("UODTTO")) <= CDate(ws_PurgeDt)) Then
      DispMsg "From/To Date Or OpenYYYYMM Cannot Be Less Than Purge Date ", etError: Exit Sub
  End If
     
  Set Rep = Nothing
  Set ADC.RepSource = Rep
  Call DispCoNm
  
    
  Call GRP_REP.Gen3LRep
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
  grpflds2 = GRP_REP.GrpFldLst(2)
  grpflds3 = GRP_REP.GrpFldLst(3)
  
      
  'DNA.15- checking if Detail section & Design code is selected
  Dim wGrp(3) As String, i As Integer, wGrpBy As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wGrpBy = 0: wNonePos = 0
  For i = 1 To 3
    If wGrp(i) = "Design Code" Then wGrpBy = i
    If wGrp(i) = "(None)" And wNonePos = 0 Then wNonePos = i
  Next i

  
  '*** To check if Detail is selected ***
  '  Dim wGrp(3) As String, i As Integer, wDetPos As Integer
  '  wGrp(1) = GRP_REP.Value(gltGroup1)
  '  wGrp(2) = GRP_REP.Value(gltGroup2)
  '  wGrp(3) = GRP_REP.Value(gltGroup3)
  '  wDetPos = 0
  '  For i = 1 To 3
  '    If wGrp(i) = "(Detail)" Then
  '      wDetPos = i: Exit For
  '    End If
  '  Next i
  '*** If Detail not selected then suppress Section 'DetDesc' Of Detail ***
  'If wDetPos = 0 Then Rep.DetDesc.Suppress = True: Rep.DETsecline.Suppress = True
  
  If ADC("UOYN2") = "Y" Then
    If CInt(Format(ADC("UoDtFr"), "mm")) > 0 And CInt(Format(ADC("UoDtFr"), "mm")) < 12 Then
        '6.1
        wTmpDt = Format$(DateAdd("d", -1, CDate(IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format$(CInt(Format(ADC("UoDtFr"), "mm")) + 1, "00") + "/01", "01/" + Format$(CInt(Format(ADC("UoDtFr"), "mm")) + 1, "00")) + "/" + CStr(CInt(Format(ADC("UoDtFr"), "yy"))))), DtShortStr())
    ElseIf CInt(Format(ADC("UoDtFr"), "mm")) = 12 Then
        '6.1
        wTmpDt = Format$(DateAdd("d", -1, CDate("01/01/" + CStr(CInt(Format(ADC("UoDtFr"), "yy")) + 1))), DtShortStr())
    End If
  End If
  Dim wValDivBy As Double
  'sv.30 when divied by amount is blank then default value is 1
  wValDivBy = IIF(ADC("UoValDivBy") = 0, 1, ADC("UoValDivBy"))
    
  '578.22 getting base currency
  Dim ws_BaseCurCd As String, wb_BaseCur As Boolean
  ws_BaseCurCd = moCn.GetFldVal("Select HBaseCurCd from Head where HCd= '" + ctSelfCmCd + "' and " + _
                  "HCoCd= '" + ADC("UoCoCdFr") + "'")
  '578.22 if currency code not given it takes base currency from head else takes from scope
  If ADC("UoCurCdFr") = "" Then
    wb_BaseCur = True
    wCurCdCnd = ""
    Rep.TxtValIn.SetText "Value In " + ws_BaseCurCd
  Else
    wb_BaseCur = False
    wCurCdCnd = " and OmCmCurCd= '" + ADC("UoCurCdFr") + "' "
    Rep.TxtValIn.SetText "Value In " + ADC("UoCurCdFr")
  End If

  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added
  wSetTmpTab = "Select 'RD' as RecTyp, space(50) as Grp1, space(50) as Grp2, space(50) as Grp3, " + _
            " space(50) as DGrp1, space(50) as DGrp2, space(50) as DGrp3, " + _
            " space(50) as G1Desc, space(50) as G2Desc, space(50) as G3Desc, " + _
            " space(1) as qFlag, OdDmCd, OdSfx, OdDmSz, OdDmStkCocd, oddmstkyy, oddmstkchr, OdDmstkno, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt as qDt,OdIdNo, " + _
            " OmPoNo as qDesc, LocPrntCd as qPrntCd, LocCd as qLoc, odcalcprc, (DDmDrQty - DDmCrQty) as qOpnQty, " + _
            " DDmDrQty as qDrQty, DDmCrQty As qCrQty, OmFrDmLoc As qFrDmLoc, OmToDmLoc As qToDmLoc, OmCmCtg, OmCmCd, OmCnvFct,OmMulDiv " + _
            " Into TmpFgStkLed From DmYyMm, OrdMst, OrdDsg, Loc, Param where 1=2 "

  moCn.CreateTmpTable "TmpFgStkLed", wSetTmpTab
  wrepcnd = ADC.RepCond
  
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd + " Option (MaxDop 1)", " Option (MaxDop 1)")
  
  wDtCndOpnYyMm = " where DYyMm < " + Format(ADC("UoDtFr"), "yyyymm")
  
  If ADC("UoOpnYyMm") <> 0 Then
    wOpnYyStr = " and DyyMm >= " + CStr(ADC("UoOpnYyMM")) + ""
  Else
    wOpnYyStr = ""
  End If
  
  '*****************  Reverse Logic *******************
  If ADC("UoYn2") = "Y" Then wDtCndOpnYyMm = " where DYyMm <= " + Format(ADC("UoDtFr"), "yyyymm")
  '****************************************************
  ADC("UoDmCdFr").CmpStr = "DDmCd>= "
  ADC("UoDmCdTo").CmpStr = "DDmCd<= "
  ADC("UoDmCdSel").CmpStr = "DDmCd in "
  ADC("UoOdSfxFr").CmpStr = "DOdSfx>= "
  ADC("UoOdSfxTo").CmpStr = "DOdSfx<= "
  ADC("UoDmSzFr").CmpStr = "DDmSz>= "
  ADC("UoDmSzTo").CmpStr = "DDmSz<= "
  ADC("UoBYyFr").CmpStr = "DDmStkYy>= "
  ADC("UoBYyTo").CmpStr = "DDmStkYy<= "
  ADC("UoBChrFr").CmpStr = "DDmStkChr>= "
  ADC("UoBChrTo").CmpStr = "DDmStkChr<= "
  ADC("UoOdDmIdFr").CmpStr = "DDmStkNo>= "
  ADC("UoOdDmIdTo").CmpStr = "DDmStkNo<= "
  ADC("UoStkCoCdFr").CmpStr = "DdmStkCoCd>= "
  ADC("UoStkCoCdTo").CmpStr = "DdmStkCoCd<= "
    
  wrepcnd = ADC.RepCond
  ' ***** Manali Trading Module
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and LocTyp= '" + adc("UoLocTypFr") + "' and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and LocTyp= 'F' and " + wrepcnd, "")
  
  '1
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added
  wsqlOpnYyMm = "Insert Into TmpFgStkLed Select 'RD' as RecTyp," + grpflds0 + ", " + _
                "'O' as qFlag, DDmCd as OdDmCd, DOdSfx as OdSfx, DDmSz as OdDmSz,DdmStkCoCd, DDmStkYy, DDmStkChr, DDmStkNo, DCoCd as OdCoCd, " + _
                "'' as OdTc, '' as OdYy, '' as OdChr, 0 as OdNo, 0 as OdSr, '' as qDt,0 as OdIdNo, '' as qDesc, " + _
                "LocPrntCd as qPrntCd, DLoc as qLoc,0.0 as OdCalcPrc, DDmDrQty- DDmCrQty as qOpnQty, 0 as qDrQty, 0 as qCrQty,  " + _
                "'' as qFrDmLoc, '' as qToDmLoc, '' as OmCmCtg, '' as OmCmCd, 1 as OmCnvFct,'' as OmMulDiv " + _
                "From DmYyMm NoLock " + _
                "Join Loc Lc On LocCoCd = DCoCd and LocCd= DLoc " + _
                "Join DsgMst on DmTcTyp='DM' and DmCd=DDmCd and DmSz='' " + _
                wDtCndOpnYyMm + wOpnYyStr + IIF(gs_Partition = ctCurrPrtn, " and DPrtKey='" + ctCurrPrtn + "' ", " and DYyMm<>'0' ") + wCnd
  moCn.Execute wsqlOpnYyMm
  
  wFrDt = "01/" + Format(ADC("UoDtFr"), "mm") + "/" + Format(ADC("UoDtFr"), "yy")
  
  wDtCndOpnOd = " where OmDt >= '" + wFrDt + "' And OmDt < '" + CStr(ADC("UoDtFr")) + "' And OmDt >= '" + wFrDt + "' And OmDt < '" + CStr(ADC("UoDtFr")) + "' "
  
  '************* Reverse Logic *************************************
  wRLQty = "(Case OmFrDmDc When 'D' Then OdOrdQty When 'C' Then -OdOrdQty Else 0 End) as qOpnQty, "
  If ADC("Uoyn2") = "Y" Then
    wDtCndOpnOd = " where OmDt >= '" & ADC("UoDtFr") & "' And OmDt <= '" & wTmpDt & "' "
    wRLQty = "(Case OmFrDmDc When 'D' Then -OdOrdQty When 'C' Then OdOrdQty Else 0 End) as qOpnQty, "
  End If
  '*****************************************************************
  
  ADC("UoDmCdFr").CmpStr = "OdDmCd>= "
  ADC("UoDmCdTo").CmpStr = "OdDmCd<= "
  ADC("UoDmCdSel").CmpStr = "OdDmCd in "
  ADC("UoOdSfxFr").CmpStr = "OdSfx>= "
  ADC("UoOdSfxTo").CmpStr = "OdSfx<= "
  ADC("UoDmSzFr").CmpStr = "OdDmSz>= "
  ADC("UoDmSzTo").CmpStr = "OdDmSz<= "
  ADC("UoBYyFr").CmpStr = "OdDmStkYy>= "
  ADC("UoBYyTo").CmpStr = "OdDmStkYy<= "
  ADC("UoBChrFr").CmpStr = "OdDmStkChr>= "
  ADC("UoBChrTo").CmpStr = "OdDmStkChr<= "
  ADC("UoOdDmIdFr").CmpStr = "OdDmStkNo>= "
  ADC("UoOdDmIdTo").CmpStr = "OdDmStkNo<= "
  ADC("UoStkCoCdFr").CmpStr = "OdDmStkCoCd>= "
  ADC("UoStkCoCdTo").CmpStr = "OdDmStkCoCd<= "
    
  wrepcnd = ADC.RepCond
  ' ***** Manali Trading Module
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and LocTyp= '" + adc("UoLocTypFr") + "' and " + wrepcnd, "")
  'wCnd = IIF(wrepcnd <> "", " and LocTyp= 'F' and " + wrepcnd, "")
 
  '2
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added, wCurCdCnd added
  wSqlOpnTdFr = "Insert Into TmpFgStkLed Select 'RD' as RecTyp," + grpflds1 + "," + _
                "'O' as qFlag, OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OmCoCd as OdCoCd, '' as OdTc,'' as OdYy, '' as OdChr, " + _
                "0 as OdNo, 0 as OdSr, '' as qDt,0 as OdIdNo, '' as qDesc, LocPrntCd as qPrntCd, OmFrDmLoc as qLoc, 0.0 as OdCalcPrc," + _
                wRLQty + " 0 as qDrQty, 0 as qCrQty, '' as qFrDmLoc, '' as qToDmLoc, OmCmCtg, OmCmCd, OmCnvFct,OmMulDiv " + _
                " From OrdMst NoLock " + _
                "Join OrdDsg on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
                " Join Loc Lc On OdCoCd= LocCoCd and OmFrDmLoc= LocCd " + _
                "Join DsgMst on DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
                wDtCndOpnOd + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " Option (MaxDop 1)"
  moCn.Execute wSqlOpnTdFr
  
  '************** Reverse Logic ****************
  wRLQty = "(Case OmToDmDc When 'D' Then OdOrdQty When 'C' Then -OdOrdQty Else 0 End) as qOpnQty, "
  If ADC("UoYn2") = "Y" Then
    wRLQty = "(Case OmToDmDc When 'D' Then -OdOrdQty When 'C' Then OdOrdQty Else 0 End) as qOpnQty, "
  End If
  '*******************************************************
  '3
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added, wCurCdCnd added
  wSqlOpnTdTo = "Insert Into TmpFgStkLed Select 'RD' as RecTyp," + grpflds1 + "," + _
                "'O' as qFlag, OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OmCoCd as OdCoCd, '' as OdTc,'' as OdYy, '' as OdChr, " + _
                "0 as OdNo, 0 as OdSr, '' as qDt,0 as OdIdNo, '' as qDesc, LocPrntCd as qPrntCd, OmToDmLoc as qLoc, 0.0 as OdCalcPrc," + _
                wRLQty + " 0 as qDrQty, 0 as qCrQty, '' as qFrDmLoc, '' as qToDmLoc, OmCmCtg, OmCmCd, OmCnvFct,OmMulDiv " + _
                "From OrdMst NoLock " + _
                "join OrdDsg on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=odPrtKey ", "") + _
                "Join Loc Lc On OdCoCd= LocCoCd and OmToDmLoc= LocCd " + _
                "Join DsgMst on DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
                wDtCndOpnOd + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey='" + ctCurrPrtn + "' ", "") + " Option (MaxDop 1)"
  moCn.Execute wSqlOpnTdTo
  
  wDtCndDet = " where OmDt between '" + CStr(ADC("UoDtFr")) + "' And '" + CStr(ADC("UoDtTo")) + "' "
  
  '4
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added, wCurCdCnd added
  wSqlDetFr = "Insert Into TmpFgStkLed Select 'RD' as RecTyp," + grpflds1 + "," + _
              "'T' as qFlag, OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt as qDt,OdIdNo, " + _
              "OmPoNo as qDesc, LocPrntCd as qPrntCd, OmFrDmLoc as qLoc, 0.0 as OdCalcPrc, 0 as qOpnQty, " + _
              "Case OmFrDmDc When 'D' Then OdOrdQty Else 0 End  as qDrQty, " + _
              "Case OmFrDmDc When 'C' Then OdOrdQty else 0 End  as qCrQty, " + _
              "(case when OmFrDmDc= 'C' then OmFrDmLoc else OmToDmLoc end) as qFrDmLoc, " + _
              "(case when OmFrDmDc= 'D' then OmFrDmLoc else OmToDmLoc end) as qToDmLoc, OmCmCtg, OmCmCd, OmCnvFct,OmMulDiv " + _
              "From OrdMst NoLock join OrdDsg on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "Join Loc Lc On LocCoCd= OdCoCd and LocCd= OmFrDmLoc " + _
              "Join DsgMst on DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
              wDtCndDet + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetFr
  
  '5
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added, wCurCdCnd added
  wSqlDetTo = "Insert Into TmpFgStkLed Select 'RD' as RecTyp," + grpflds1 + "," + _
              "'T' as qFlag, OdDmCd, OdSfx, OdDmSz,OdDmStkCoCd, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr, OmDt as qDt,OdIdNo, " + _
              "OmPoNo  as qDesc, LocPrntCd as qPrntCd, OmToDmLoc as qLoc, 0.0 as OdCalcPrc, 0 as qOpnQty, " + _
              "Case OmToDmDc When 'D' Then OdOrdQty Else 0 End  as qDrQty, " + _
              "Case OmToDmDc When 'C' Then OdOrdQty else 0 End  as qCrQty, " + _
              "(case when OmFrDmDc= 'C' then OmFrDmLoc else OmToDmLoc end) as qFrDmLoc, " + _
              "(case when OmFrDmDc= 'D' then OmFrDmLoc else OmToDmLoc end) as qToDmLoc, OmCmCtg, OmCmCd, OmCnvFct,OmMulDiv " + _
              "From OrdMst NoLock join OrdDsg on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
              "Join Loc Lc On LocCoCd= OdCoCd and LocCd= OmToDmLoc " + _
              "Join DsgMst on DmTcTyp='DM' and DmCd=OdDmCd and DmSz='' " + _
              wDtCndDet + wCnd + wCurCdCnd + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey='" + ctCurrPrtn + "' ", "") + " Option (MaxDop 1)"
  moCn.Execute wSqlDetTo
  'Vk.21 -OmCmCtg, OmCmCd added
  '578.22 OmCnvFct,OmMulDiv added
  wSqlFinal = "Insert into TmpFgStkLed Select 'F' as RecTyp," + _
      "max(Grp1) as Grp1, max(Grp2) as Grp2, max(Grp3) as Grp3, " + _
      "DGrp1, DGrp2, DGrp3, max(G1Desc) as G1Desc, max(G2Desc) as G2Desc, max(G3Desc) as G3Desc, " + _
      "max(qFlag) as qFlag, Max(OdDmCd) as OdDmCd, Max(OdSfx) as OdSfx, max(OdDmSz) as OdDmSz,max(OdDmStkCoCd) as OdDmStkCoCd, Max(OdDmStkYy) as OdDmStkYy, Max(OdDmStkChr) as OdDmStkChr, Max(OdDmStkNo) as OdDmStkNo, " + _
      "OdCoCd, OdTc, OdYy, OdChr, OdNo, max(OdSr) as OdSr, max(qDt),max(OdIdNo), Max(qDesc), " + _
      "max(qPrntCd) as qPrntCd, max(qLoc) as qLoc, 0.0 as OdCalcPrc, sum(qOpnQty) as qOpnQty,  " + _
      "sum(qDrQty) as qDrQty, sum(qCrQty) as qCrQty, " + _
      "max(qFrDmLoc) as qFrDmLoc, max(qToDmLoc) as qToDmLoc, max(OmCmCtg) as OmCmCtg, max(OmCmCd) as OmCmCd, max(OmCnvFct) as OmCnvFct, max(OmMulDiv) as OmMulDiv " + _
      "From TmpFgStkLed Tmp " + _
      "Group By DGrp1, DGrp2, DGrp3, OdDmCd, OdSfx, OdDmSz, OdDmStkYy, OdDmStkChr, OdDmStkNo, OdCoCd, OdTc, OdYy, OdChr, OdNo , OdSr"
      
      
  moCn.Execute wSqlFinal
  'Vk.21 -OmCmCtg, OmCmCd added
  'sv.30 wValDivBy added
  '578.22 OmCnvFct,OmMulDiv added
  ADC.RepRecSource = "Select Grp1 as Grp1, Grp2, Grp3," + _
      "DGrp1, DGrp2, DGrp3, G1Desc as G1Desc, G2Desc as G2Desc, G3Desc as G3Desc, " + _
      "qFlag as qFlag, Tmp.OdDmCd as OdDmCd, Tmp.OdSfx as OdSfx, Tmp.OdDmSz as OdDmSz,Tmp.OdDmStkCoCd as OdDmStkCoCd, Tmp.OdDmStkYy as OdDmStkYy, Tmp.OdDmStkChr as OdDmStkChr,Tmp.OdDmStkNo as OdDmStkNo, " + _
      "OdCoCd, OdTc, OdYy, OdChr, OdNo, OdSr as OdSr, qDt,OdIdNo, qDesc, " + _
      "qPrntCd as qPrntCd, qLoc as qLoc, isnull(" + IIF(wb_BaseCur, "(Case when OmMulDiv= 'M' then OdForVal.OdSalPrc/Tmp.OmCnvFct else OdForVal.OdSalPrc*Tmp.OmCnvFct end) ", "OdForVal.OdSalPrc") + ",0)/" + CStr(wValDivBy) + " as OdCalcPrc, qOpnQty as qOpnQty,  " + _
      "qDrQty as qDrQty, qCrQty, qFrDmLoc as qFrDmLoc, qToDmLoc as qToDmLoc, Tmp.OmCmCtg, Tmp.OmCmCd " + _
         " from TmpFgStkLed Tmp " + _
         " left outer join " + _
         " (Select OrdDsg.OdDmCd, OrdDsg.OdSfx, OrdDsg.OdDmSz, (OrdDsg.OdDmStkYy) as OdDmStkYy, (OrdDsg.OdDmStkChr) as OdDmStkChr, (OrdDsg.OdDmStkNo) as OdDmStkNo,OrdDsg.OdSalPrc from " + _
         " OrdDsg join (Select OdDmCd, OdSfx, OdDmSz, max(OdDmStkYy) as OdDmStkYy, max(OdDmStkChr) as OdDmStkChr, max(OdDmStkNo) as OdDmStkNo, min(OdIdNo) as OdIdNo from OrdDsg " + _
                      " join OrdMst on OrdDsg.OdYy = OrdMst.OmYy and OrdDsg.OdTc = OrdMst.OmTc and OrdDsg.OdChr = OrdMst.OmChr and OrdDsg.OdNo = OrdMst.OmNo" + _
                      " left outer join Loc FrLocMst on FrLocMst.LocCd = OrdMst.OmFrDmLoc " + _
                      " left outer join Loc ToLocMst on ToLocMst.LocCd = OrdMst.OmToDmLoc " + _
                      "  where ((isnull(FrLocMst.LocTyp,'')  = 'S' and OrdMst.OmFrDmDc = 'D') or (isnull(ToLocMst.LocTyp,'')  = 'S' and OmToDmDc = 'D')) " + _
                      "  and exists(Select 'x' from Param Where PTyp='TC' and PMCd = OdTc and PValue='" + ctJtTcTyp + "') " + _
                      " Group By " + _
                      " OdDmCd, OdSfx, OdDmSz, case when OdDmStkNo <>'0' then OdDmStkYy + OdDmStkChr+str(OdDmStkNo) else '' end) OdMinRecForVal       " + _
                      " on OrdDsg.OdIdNo = OdMinRecForVal.OdIdNo) OdForVal " + _
                      " on Tmp.OdDmCd = OdForVal.OdDmCd and Tmp.OdSfx = OdForVal.OdSfx and Tmp.OdDmSz = OdForVal.OdDmSz and Tmp.OdDmStkYy =OdForVal.OdDmStkYy" + _
                      " and Tmp.OdDmStkChr =OdForVal.OdDmStkChr and Tmp.OdDmStkNo = OdForVal.OdDmStkNo " + _
                      "join DsgMst on DmTcTyp= 'DM' and DmCd= Tmp.OdDmCd and DmSz= '' " + _
                      "And (Select Count(*) From UsrOptDsgAna Join DsgAna on " + IIF(gs_Partition = ctCurrPrtn, " DaPrtKey='" + ctCurrPrtn + "' and ", "") + " DaAnaSr = UdAnaSr " + _
                      "where (DaAnaCd between UdAnaCdFr and UdAnaCdTo) and DaTcTyp=DmTcTyp and DaCd= DmCd and DaSz= DmSz " + _
                      "and UdUsrCd = '" + ADC.UsrCd + "' and UdMnuCd= '" + ADC.MenuCd + "')= " & (GrdDsgAna.Rows - 1) & " " + _
                      " where RecTyp = 'F'"

  'DNA.15-If Show picture is not selected then will not show the picture
  'if selected then show the picture in footer3 / header1 or 2
  If ADC("UoYN1") = "Y" And wGrpBy <> 0 Then
    If wGrpBy = 1 And wNonePos = 2 Then
      Rep.PicSecF.Suppress = False: Rep.PicSecH2.Suppress = True:   Rep.PicSecH1.Suppress = True
    ElseIf wGrpBy = 1 And wNonePos = 3 Then
      Rep.PicSecF.Suppress = True:  Rep.PicSecH2.Suppress = False:  Rep.PicSecH1.Suppress = True
    ElseIf wGrpBy = 1 And wNonePos = 0 Then
      Rep.PicSecF.Suppress = True:  Rep.PicSecH2.Suppress = True:   Rep.PicSecH1.Suppress = False
    ElseIf wGrpBy = 2 And wNonePos = 3 Then
      Rep.PicSecF.Suppress = False: Rep.PicSecH2.Suppress = True:   Rep.PicSecH1.Suppress = True
    ElseIf wGrpBy = 2 And wNonePos = 0 Then
      Rep.PicSecF.Suppress = True: Rep.PicSecH2.Suppress = False:   Rep.PicSecH1.Suppress = True
    ElseIf wGrpBy = 3 And wNonePos = 0 Then
      Rep.PicSecF.Suppress = False: Rep.PicSecH2.Suppress = True:   Rep.PicSecH1.Suppress = True
    End If
  Else
    Rep.PicSecF.Suppress = True:    Rep.PicSecH2.Suppress = True:   Rep.PicSecH1.Suppress = True
  End If
    
  Rep.TxtDtFr.SetText ADC("UoDtFr")
  Rep.TxtDtTo.SetText ADC("UoDtTo")
  Rep.txtLocTypFr.SetText ADC("UoLocTypFr")
  Rep.TxtLocTypTo.SetText ADC("UoLocTypTo")
  Rep.TxtLocTypSel.SetText ADC("UoLocTypSel")
  'Rep.txtLocTyp.SetText adc("UoLocTypFr")     ' ***** Manali Trading Module
  
  Rep.RecordSelectionFormula = " {rdo.qDrQty} <> 0 Or {rdo.qCrQty} <> 0 Or {rdo.qOpnQty} <> 0 "
  
  If UCase(ADC.MenuCd) = UCase("RepJtStkLed") Then
    Rep.TxtHead.SetText "Stock Ledger (Jwlry)"
  End If
  'sv.30 passing divide by value to dsr
  GRP_REP.SetFormula Rep, "wDivByVal", ADC("UoValDivBy")
  
  'sv.30 when divied by amount>1 then follwing text will display in header
  If ADC("UoValDivBy") > 1 Then Rep.TxtDivByVal.SetText "(Values in " + Format(ADC("UoValDivBy"), "##,##,##0") + ")"
  
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(ADC)
  
Exit Sub

errhld:
    moCn.Execute "Delete From TmpFgStkLed"
    DispMsg Err.Description, etError
    ADC.RepRecSource = "Select * From TmpFgStkLed"
  
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  'DNA.15-If check box is checked then Show picture is Yes else No
  Case Is = 1
    If .Value = Checked Then ADC("UoYN1") = "Y" Else ADC("UoYN1") = "N"
  Case Is = 2     ' Tag = "ShowLine"
    If .Value = Checked Then ADC("UoYN2") = "Y" Else ADC("UoYN2") = "N"
  End Select
End With
End Sub

Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
'pg.8- Set visible /invisible state for analysis frame
Private Sub CmdSc_Click()
  If FraSc.Visible = False Then
    If CRV_REP.Visible = True Then Exit Sub
    FraSc.Enabled = True
    FraSc.Visible = True
    TAB_REP.Enabled = False
    ADC.AllowSave = False
    FraSc.ZOrder
  Else
    TAB_REP.Enabled = True
    ADC.AllowSave = True
    FraSc.Enabled = False
    FraSc.Visible = False
  End If
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
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
'pg
  ADC.Child = "GrdDsgAna"
  Set GrdDsgAna.PrevCtl = ADC
  Set GrdDsgAna.NextCtl = ADC
  GrdDsgAna.HotKey = "A"
  
  gb_CoCdFor = False
  Call InitProp(Me)
  If ws_MultiCoMod = True Then
    Set ADC.FirNKeyCtl = ADC("UoCoCdFr")
    ADC("UOCOCDFR").Enabled = True: ADC("UOCOCDTO").Visible = True: ADC("UOCOCDSEL").Visible = True
    LblPrntCd.Visible = True: ADC("UOPRNTCDFR").Visible = True: ADC("UOPRNTCDTO").Visible = True: ADC("UOPRNTCDSEL").Visible = True
  Else
    Set ADC.FirNKeyCtl = ADC("UoLocTypFr")
    ADC("UOCOCDFR").Enabled = False: ADC("UOCOCDTO").Visible = False: ADC("UOCOCDSEL").Visible = False
    LblPrntCd.Visible = False: ADC("UOPRNTCDFR").Visible = False: ADC("UOPRNTCDTO").Visible = False: ADC("UOPRNTCDSEL").Visible = False
  End If
  
  ' ***** Manali Trading Module
  'gs_LocTyp = "'F'"
  
  Call SetGroupSort
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Dim ws_NewVal As String
  
  Select Case UCase(IdName)
  ' ***** Manali Trading Module
  Case Is = UCase("UoLocTypFr")
    ws_NewVal = pv_NewValue
    gs_LocTyp = IIF(ws_NewVal = "", "", "'" + ws_NewVal + "'")
  'sv.30 divided by amount cannot be 0
  Case Is = UCase("UoValDivBy")
    If pv_NewValue = 0 Then Cancel = True: ErrMsg = "Divided By Amount Cannot be Zero": Exit Sub
    
  End Select

  gb_CoCdFor = False
  
  ' ***** Manali Trading Module
  'gs_LocTyp = "'F'"
  
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  ' ***** Manali Trading Module
  Select Case UCase(IdName)
  Case Is = UCase$("UoIssLocFr")
    gs_LocTyp = IIF(ADC("UoLocTypFr") = "", "", "'" + ADC("UoLocTypFr") + "'")
  Case Is = UCase$("UoIssLocTo")
    gs_LocTyp = IIF(ADC("UoLocTypTo") = "", "", "'" + ADC("UoLocTypTo") + "'")
  Case Is = UCase$("UoIssLocSel")
    gs_LocTyp = IIF(ADC("UoLocTypFr") = "", "", "'" + ADC("UoLocTypFr") + "'")
  
  End Select
  'gs_LocTyp = "'F'"
  SetProp Me, IdName, When
End Sub


Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  ADC("UOUSRCD") = ADC.UsrCd
  ADC("UOMNUCD") = ADC.MenuCd
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  If ws_MultiCoMod = False Then _
      ADC("UoCoCdFr") = gs_CoCd: ADC("UoCoCdTo") = gs_CoCd: ADC("UoCoCdSel") = "": ADC("UoCoCdFr").Enabled = False
  'DNA.15-Setting default values to Show picture option
  ADC("UoYN1") = "N"
  ADC("UoOpnYyMm") = 0
  ' ***** Manali Trading Module
  gs_LocTyp = "'" + ADC("UoLocTypFr") + "'"
  'gs_LocTyp = "'F'"
  
  ADC("UOYN2") = "Y"
  If ADC("UoYN2") = "Y" Then ChkBoxArr(2).Value = Checked Else ChkBoxArr(2).Value = Unchecked  ' ***** Manali 3.7.0
  
  'DNA.15- Setting Check box values as Show picture Yes / No
  If ADC("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  ADC("UoValDivBy") = 1     'sv.30 setting default value
  'sv.30 divide by value will be visible only for jemr stock ledger report
  If UCase(ADC.MenuCd) <> UCase("RepJtStkLed") Then
    LblValDivBy.Visible = False
    ADC("UoValDivBy").Visible = False
    LblCurCd.Visible = False:   ADC("UOCURCDFR").Visible = False    '578.22 Currency code will be visible only for jemr stock ledger report
  End If
  
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + IIF(ws_MultiCoMod = True, gs_CoCd, ADC("UOCOCDFR")) + "' and " + _
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
  
  
  If ws_MultiCoMod = True Then
      Rep.TxtCoCdToLbl.Suppress = False: Rep.TxtCoCdSelLbl.Suppress = False
      Rep.TxtCoCdTo.SetText ADC("UOCOCDTO"): Rep.TxtCoCdSel.SetText ADC("UOCOCDSEL")
  Else
      Rep.TxtCoCdTo.Suppress = True: Rep.TxtCoCdSel.Suppress = True
      Rep.TxtCoCdToLbl.Suppress = True: Rep.TxtCoCdSelLbl.Suppress = True
  End If
  Rep.TxtCoCdFr.SetText ADC("UOCOCDFR")
  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
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

'pg.8-checking Valid Analysis Sr & field
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
      .Store "UdUsrCd", ADC.UsrCd
      .Store "UdMnuCd", ADC.MenuCd
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
  GrdDsgAna.RecSource = "Select * from UsrOptDsgAna where UdUsrCd= '" + ADC.UsrCd + "' and " + _
                        "UdMnuCd= '" + ADC.MenuCd + "'"
                        
End Sub

Private Sub GrdDsgAna_Validate(Cancel As Boolean)
  Cancel = GrdDsgAna.Validate
End Sub

