VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpNetStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Stock Report"
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
      Left            =   11295
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   9630
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4455
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   5295
      TabIndex        =   70
      Top             =   9615
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   71
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
         TabIndex        =   72
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
      TabIndex        =   62
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
      Left            =   6510
      TabIndex        =   69
      Top             =   9585
      Visible         =   0   'False
      Width           =   735
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   120
         TabIndex        =   66
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
      TabIndex        =   65
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   64
         Top             =   60
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpNetStk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpNetStk.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   68
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   63
               Top             =   360
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   67
            Top             =   360
            Width           =   15105
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Zero Req"
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
               Left            =   0
               TabIndex        =   60
               Tag             =   "ShowDet"
               ToolTipText     =   "Check To Show Zero Req"
               Top             =   6840
               Width           =   1800
            End
            Begin VB.Frame FraOptRpt 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   1480
               TabIndex        =   106
               ToolTipText     =   "Select Rm Code / Exact By ALL Size / Upto Range Size"
               Top             =   6240
               Width           =   6090
               Begin VB.OptionButton OptRpt 
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
                  Height          =   285
                  Index           =   2
                  Left            =   3720
                  TabIndex        =   59
                  ToolTipText     =   "Select Rm Code / Exact By ALL Size / Upto Range Size"
                  Top             =   0
                  Width           =   1965
               End
               Begin VB.OptionButton OptRpt 
                  Caption         =   "Exact By ALL Size"
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
                  Left            =   1440
                  TabIndex        =   58
                  ToolTipText     =   "Select Rm Code / Exact By ALL Size / Upto Range Size"
                  Top             =   0
                  Width           =   2085
               End
               Begin VB.OptionButton OptRpt 
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
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   57
                  ToolTipText     =   "Select Rm Code / Exact By ALL Size / Upto Range Size"
                  Top             =   0
                  Width           =   1305
               End
            End
            Begin VB.Frame FraOptQW 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   1485
               TabIndex        =   53
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   5880
               Width           =   2235
               Begin VB.OptionButton OptQW 
                  Caption         =   "Wt"
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
                  TabIndex        =   54
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   765
               End
               Begin VB.OptionButton OptQW 
                  Caption         =   "Qty"
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
                  Left            =   1320
                  TabIndex        =   55
                  ToolTipText     =   "Select either Quantity Or Weight"
                  Top             =   0
                  Width           =   825
               End
            End
            Begin VB.Frame FraOptRm 
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   1500
               TabIndex        =   100
               ToolTipText     =   "Select either Quantity Or Weight"
               Top             =   5280
               Width           =   1305
               Begin VB.OptionButton OptRm 
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
                  Index           =   0
                  Left            =   0
                  TabIndex        =   51
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.OptionButton OptRm 
                  Caption         =   "Order"
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
                  TabIndex        =   52
                  Top             =   285
                  Width           =   1365
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Detail    "
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
               TabIndex        =   49
               Tag             =   "ShowDet"
               ToolTipText     =   "Check To Show Location And order Details"
               Top             =   4995
               Width           =   1680
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4140
               TabIndex        =   5
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1215
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
               Index           =   15
               Left            =   6840
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1215
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4140
               TabIndex        =   3
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   930
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
               Index           =   3
               Left            =   2265
               TabIndex        =   1
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   645
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1485
               TabIndex        =   2
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   931
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
               Index           =   7
               Left            =   1485
               TabIndex        =   4
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1215
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
               Index           =   28
               Left            =   4140
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Category"
               Top             =   2070
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGTO"
               IdName          =   "UODMCTGTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   6840
               TabIndex        =   20
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2070
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UODMCTGSEL"
               IdName          =   "UODMCTGSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1485
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Category"
               Top             =   2070
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   5
               DataType        =   4
               DataField       =   "UODMCTGFR"
               IdName          =   "UODMCTGFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   31
               Left            =   4140
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Code"
               Top             =   2355
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDTO"
               IdName          =   "UODMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   32
               Left            =   6840
               TabIndex        =   23
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2355
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1485
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Code"
               Top             =   2355
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   15
               DataField       =   "UODMCDFR"
               IdName          =   "UODMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   40
               Left            =   4140
               TabIndex        =   25
               ToolTipText     =   "Enter To Customer"
               Top             =   2640
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDTO"
               IdName          =   "UOCMCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   41
               Left            =   6840
               TabIndex        =   26
               ToolTipText     =   "Enter Customer Selection"
               Top             =   2640
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1485
               TabIndex        =   24
               ToolTipText     =   "Enter From Customer"
               Top             =   2647
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOCMCDFR"
               IdName          =   "UOCMCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   4140
               TabIndex        =   28
               ToolTipText     =   "Enter To Purchase Order "
               Top             =   2925
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOTO"
               IdName          =   "UOPONOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   44
               Left            =   6840
               TabIndex        =   29
               ToolTipText     =   "Enter Purchase Order Selection"
               Top             =   2925
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPONOSEL"
               IdName          =   "UOPONOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   1485
               TabIndex        =   27
               ToolTipText     =   "Enter From Purchase Order "
               Top             =   2925
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOPONOFR"
               IdName          =   "UOPONOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4140
               TabIndex        =   36
               ToolTipText     =   "Enter To Order Date"
               Top             =   3495
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
               Left            =   1485
               TabIndex        =   35
               ToolTipText     =   "Enter From Order Date"
               Top             =   3495
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
               Index           =   49
               Left            =   4140
               TabIndex        =   38
               ToolTipText     =   "Enter To Prd Delivery Date"
               Top             =   3780
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   1485
               TabIndex        =   37
               ToolTipText     =   "Enter From Prd Delivery Date"
               Top             =   3780
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   4905
               TabIndex        =   11
               ToolTipText     =   "Enter To Order Year"
               Top             =   1500
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYTO"
               IdName          =   "UOOMYYTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2280
               TabIndex        =   8
               ToolTipText     =   "Enter From Order Year"
               Top             =   1500
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOOMYYFR"
               IdName          =   "UOOMYYFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   10
               Left            =   1485
               TabIndex        =   7
               ToolTipText     =   "Enter From Order Tc"
               Top             =   1500
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   2685
               TabIndex        =   9
               ToolTipText     =   "Enter From Order Character"
               Top             =   1500
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRFR"
               IdName          =   "UOOMCHRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   4140
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Tc"
               Top             =   1500
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5325
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Character"
               Top             =   1500
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOOMCHRTO"
               IdName          =   "UOOMCHRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   5025
               TabIndex        =   16
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1785
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   52
               Left            =   2385
               TabIndex        =   14
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1785
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   53
               Left            =   1485
               TabIndex        =   13
               ToolTipText     =   "Enter From Order Number"
               Top             =   1785
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOFR"
               IdName          =   "UOOMNOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   54
               Left            =   4140
               TabIndex        =   15
               ToolTipText     =   "Enter To Order Number"
               Top             =   1785
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOOMNOTO"
               IdName          =   "UOOMNOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1485
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
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   1485
               TabIndex        =   39
               ToolTipText     =   "Enter From  Priority Code"
               Top             =   4065
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   56
               Left            =   4140
               TabIndex        =   40
               ToolTipText     =   "Enter To  Priority Code"
               Top             =   4065
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   57
               Left            =   6840
               TabIndex        =   41
               ToolTipText     =   "Enter  Priority Code Selection"
               Top             =   4065
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4140
               TabIndex        =   43
               ToolTipText     =   "Enter To FrBag Location"
               Top             =   4350
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   5
               DataField       =   "UOISSLOCTO"
               IdName          =   "UOISSLOCTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   6840
               TabIndex        =   44
               ToolTipText     =   "Enter FrBag Location Selection"
               Top             =   4350
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOISSLOCSEL"
               IdName          =   "UOISSLOCSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1485
               TabIndex        =   42
               ToolTipText     =   "Enter From FrBag Location"
               Top             =   4350
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOISSLOCFR"
               IdName          =   "UOISSLOCFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   24
               Left            =   1905
               TabIndex        =   48
               ToolTipText     =   "Show Location And Order Detail ? (Enter 'Y'/ 'N')"
               Top             =   4995
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
               Index           =   2
               Left            =   6840
               TabIndex        =   97
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   930
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   16
               Left            =   3075
               TabIndex        =   50
               ToolTipText     =   "Show All Or Order Specific RM (Enter 'Y' For All Or 'N' For Order Specific RM)"
               Top             =   5280
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
               Index           =   4
               Left            =   2775
               TabIndex        =   31
               ToolTipText     =   "Enter From Sub Purchase Order"
               Top             =   3210
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOFR"
               IdName          =   "UOODPONOFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   5430
               TabIndex        =   33
               ToolTipText     =   "Enter To Sub Purchase Order "
               Top             =   3210
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOODPONOTO"
               IdName          =   "UOODPONOTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   6840
               TabIndex        =   34
               ToolTipText     =   "Enter Sub Purchase Order Selection"
               Top             =   3210
               Visible         =   0   'False
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOODPONOSEL"
               IdName          =   "UOODPONOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   3720
               TabIndex        =   102
               ToolTipText     =   "Enter 'Y' for Quantity and 'N' for Weight"
               Top             =   5880
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN3"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOYN3"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1485
               TabIndex        =   30
               ToolTipText     =   "Enter From Sub Customer"
               Top             =   3210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTFR"
               IdName          =   "UOODSUBCUSTFR"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   4140
               TabIndex        =   32
               ToolTipText     =   "Enter To Sub Customer"
               Top             =   3210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODSUBCUSTTO"
               IdName          =   "UOODSUBCUSTTO"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   86
               Left            =   6840
               TabIndex        =   17
               ToolTipText     =   "Enter Order Number Selection"
               Top             =   1785
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOOMNOSEL"
               IdName          =   "UOOMNOSEL"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   72
               Left            =   1485
               TabIndex        =   45
               ToolTipText     =   "Enter Raw Material Stone Class From "
               Top             =   4635
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
               Index           =   74
               Left            =   4140
               TabIndex        =   46
               ToolTipText     =   "Enter Raw Material Stone Class To "
               Top             =   4635
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
               Index           =   76
               Left            =   6840
               TabIndex        =   47
               ToolTipText     =   "Enter Raw Material Stone Class Selection"
               Top             =   4635
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSTNCLSSEL"
               IdName          =   "UOSTNCLSSEL"
               CmpStr          =   "ISNULL(SC.PSCd,'') In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   7680
               TabIndex        =   108
               ToolTipText     =   "Enter 'Y' for Quantity and 'N' for Weight"
               Top             =   6240
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOYN2"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   21
               Left            =   2160
               TabIndex        =   109
               ToolTipText     =   "Enter 'Y' for Quantity and 'N' for Weight"
               Top             =   6840
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN4"
               ReCalcParent    =   "UOYN"
               IdName          =   "UOYN4"
            End
            Begin VB.Label LblOptRpt 
               BackStyle       =   0  'Transparent
               Caption         =   "Requirement By"
               BeginProperty Font 
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
               Left            =   30
               TabIndex        =   107
               Top             =   6240
               Width           =   1365
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
               Left            =   30
               TabIndex        =   105
               Top             =   4635
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
               Index           =   7
               Left            =   5370
               TabIndex        =   104
               Top             =   3240
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
               Index           =   16
               Left            =   2715
               TabIndex        =   103
               Top             =   3240
               Width           =   105
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Wt/Qty"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   0
               TabIndex        =   101
               Top             =   5880
               Width           =   1455
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
               Index           =   5
               Left            =   30
               TabIndex        =   99
               Top             =   3210
               Width           =   1425
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Rm Details"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   570
               Left            =   30
               TabIndex        =   98
               Top             =   5280
               Width           =   1485
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Detail"
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
               Left            =   30
               TabIndex        =   96
               Top             =   4995
               Visible         =   0   'False
               Width           =   1485
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
               Index           =   6
               Left            =   30
               TabIndex        =   95
               Top             =   4350
               Width           =   1485
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
               Height          =   285
               Left            =   30
               TabIndex        =   94
               Top             =   4065
               Width           =   1485
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
               TabIndex        =   93
               Top             =   660
               Width           =   105
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
               TabIndex        =   92
               Top             =   1500
               Width           =   1485
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
               TabIndex        =   91
               Top             =   1785
               Width           =   1485
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
               TabIndex        =   90
               Top             =   1500
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
               TabIndex        =   89
               Top             =   1500
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
               TabIndex        =   88
               Top             =   1500
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
               TabIndex        =   87
               Top             =   1500
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
               TabIndex        =   86
               Top             =   1785
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
               TabIndex        =   85
               Top             =   1785
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
               Left            =   30
               TabIndex        =   84
               Top             =   3780
               Width           =   1485
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
               TabIndex        =   83
               Top             =   3495
               Width           =   1485
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
               TabIndex        =   82
               Top             =   2925
               Width           =   1485
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
               TabIndex        =   81
               Top             =   2640
               Width           =   1485
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
               TabIndex        =   80
               Top             =   2355
               Width           =   1485
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
               TabIndex        =   79
               Top             =   2070
               Width           =   1485
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
               TabIndex        =   78
               ToolTipText     =   "Location"
               Top             =   645
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
               TabIndex        =   77
               Top             =   930
               Width           =   1485
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
               TabIndex        =   76
               Top             =   1215
               Width           =   1485
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
               TabIndex        =   75
               Top             =   135
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
               TabIndex        =   74
               Top             =   135
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
               TabIndex        =   73
               Top             =   135
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpNetStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepNetStk
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06

Private Sub SetGroupSort()
  GRP_REP.Add "Rm Code", Array("Rtrim(max(RmSCtg))+Rtrim(max(RmCd))", "Rtrim(OrRmSCtg)+Rtrim(OrRmCd)"), Array("'('+ max(RmSCtg)+ ') '+ max(RmCd)", "'('+ OrRmSCtg+ ') '+ OrRmCd"), "RmSCtg,RmCd", "hRmSCtg,hRmCd", "Rm SubCtg,Parent Cd", "", "", Array("SubString(max(RmDesc), 1, 30) ", "SubString(RmDesc, 1, 30) ")
  GRP_REP.Add "Rm SubCtg", Array("max(RmSCtg)", "OrRmSCtg"), Array("max(RmSCtg)", "OrRmSCtg"), "RmSCtg", "hRmSCtg", "", "", "", Array("(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = max(RmCtg) and PSCd= max(RmSCtg)) ", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = OrRmCtg and PSCd= OrRmSCtg) ")
  GRP_REP.Add "Parent Cd", Array("Rtrim(max(RmSCtg))+(case when max(RmParentCd)= '' then max(RmCd) else max(RmParentCd) end)", "Rtrim(OrRmSCtg)+(case when RmParentCd= '' then OrRmCd else RmParentCd end)"), Array("'('+ max(RmSCtg)+ ') '+ (case when max(RmParentCd)= '' then max(RmCd) else max(RmParentCd) end)", "'('+ OrRmSCtg+ ') '+ (case when RmParentCd= '' then OrRmCd else RmParentCd end)"), "OrRmSCtg", "hOrRmSCtg", "Rm SubCtg", "", "", Array("(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= max(RmParentCd)) ", "(Select PDesc from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd) ")
  GRP_REP.Add "Range Size", Array("max(RmSCtg)+max(IsNull(RrRngSz, ''))", "RmSCtg+IsNull(RrRngSz, '')"), Array("'('+ max(RmSCtg)+ ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ ') '+IsNull(RrRngSz, '')"), "RmSCtg,qRngSz", "hRmSCtg,hqRngSz", "Rm SubCtg", "", "", ""
  GRP_REP.Add "RmCd/RangeSz", Array("max(RmSCtg)+max(IsNull(RrRngSz, ''))+max(RmCd)", "RmSCtg+IsNull(RrRngSz, '')+RmCd"), Array("'('+ max(RmSCtg)+ '/' + max(RmCd) + ') '+max(IsNull(RrRngSz, ''))", "'('+ RmSCtg+ '/' + RmCd + ') '+IsNull(RrRngSz, '')"), "RmSCtg,RmCd,qRngSz", "hRmSCtg,hRmCd,hqRngSz", "Rm SubCtg,Rm Code,Parent Cd,Range Size", "", "", ""
  '6.1-98 Stone Class group sort added
  GRP_REP.Add "Stone Class", Array("max(RmCtg+IsNull(SC.PSCd, ''))", "RmCtg+IsNull(SC.PSCd, '')"), Array("'('+ RTrim(max(RmCtg))+') '+max(IsNull(SC.PSCd, ''))", "'('+ RTrim(RmCtg)+') '+IsNull(SC.PSCd, '')"), "", "", "", "", "", Array("max(SC.PDesc)", "SC.PDesc")
  '6.1B-43
  GRP_REP.Add "Rm Sz", Array(GetRmSzDBFldNm("max(SRmSz)", "max(SRmSz2)", "max(SRmSz3)"), GetRmSzDBFldNm("(OrLn1)", "(OrLn2)", "(OrLn3)")), Array(GetRmSzDispFldNm("max(SRmSz)", "max(SRmSz2)", "max(SRmSz3)"), GetRmSzDispFldNm("(OrLn1)", "(OrLn2)", "(OrLn3)")), "wRmSz", "hqStkSz", "Rm SubCtg,Rm Code,Range Size,RmCd/RangeSz", "", "", ""

End Sub
Private Sub adc_setreprecsource()
  ' ***** Manali 3.03 - RisRmSzDesc added and join with RmIdSz added
  '*** Report Sql ***
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  SLoc, SRmCd as RmCd, OrRmCtg as RmCtg, OrRmSCtg as RmSCtg,
  '  SRmSz as qStkSz, RisRmIdSz as qRmSzDesc, RrRngSz as qRngSz, space(1) as qSzYN,
  '  SRmDrWt as qAvlblWt, SRmDrWt as qOrRmWt, OdPrdQty as qBalPrdQty,
  '  SRmDrWt As qRmPureWt, SRmDrWt As qOrRmPureWt,
  '  OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo
  '  From SYyMm, OrdRm, OrdDsg, RmRt, RmIdSz
  '  Where 1 = 2d
  '*** Report Sql ***

  Dim grpflds0 As String, grpflds1 As String
  Dim ws_RepCnd As String, ws_Cnd As String, ws_SqlStrg As String
  
  Dim wd_BaseGldRmPurity As Double, wd_BasePltRmPurity As Double
  
  ' ****** Sachin 2.12 - 02/12/05 - SJM
  Dim wd_BaseSilRmPurity As Double
  ' ****** Sachin 2.12 - 02/12/05 - SJM
  
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  Dim wd_BasePdRmPurity As Double
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  
  Dim ws_HDTPrdPtr As String
  
  '****** Manali 3.03 - 09/07/08
  Dim wRmSzDesc As String
  '****** Manali 3.03 - 09/07/08
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in where clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  
  grpflds0 = GRP_REP.GrpFldLst(0)
  grpflds1 = GRP_REP.GrpFldLst(1)
    
  '*** To check if Detail is selected ***
  Dim ws_Grp(6) As String, wi_i As Integer, wi_DetPos As Integer
  ws_Grp(1) = GRP_REP.Value(gltGroup1)
  ws_Grp(2) = GRP_REP.Value(gltGroup2)
  ws_Grp(3) = GRP_REP.Value(gltGroup3)
  wi_DetPos = 0
  For wi_i = 1 To 6
    If ws_Grp(wi_i) = "(Detail)" Then
      wi_DetPos = wi_i: Exit For
    End If
  Next wi_i
    
  If wi_DetPos = 0 Or adc("UoYn") = "N" Then
    Rep.BoxHdLoc.Suppress = True: Rep.BoxDetLoc.Suppress = True: Rep.BoxFtLoc.Suppress = True
    Rep.BoxHdOrd.Suppress = True: Rep.BoxDetOrd.Suppress = True: Rep.BoxFtOrd.Suppress = True
    If wi_DetPos = 0 Then Rep.G4F.Suppress = True
  End If
    
  wd_BaseGldRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'G'), 1) ")
  wd_BasePltRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'P'), 1) ")
  ' ****** Sachin 2.12 - 02/12/05 - SJM
  wd_BaseSilRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'S'), 1) ")
  ' ****** Sachin 2.12 - 02/12/05 - SJM
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  wd_BasePdRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'L'), 1) ")
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L'
  ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + adc("UoCoCdFr") + "'"))

  '****** Manali 3.03 - 09/07/08
  wRmSzDesc = "case when (Select PDesc225 from Param where PTyp= 'RMSCTG' and PMCd= IsNull(max(RisRmCtg), '') and PSCd= max(RmSCtg))= 'Y' then IsNull(max(RisRmSzDesc), '') Else '' End "
  '****** Manali 3.03 - 09/07/08
  
''  ws_RepCnd = adc.RepCond
''  ws_Cnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")

  '**************** Logic For Getting qRngSz fields ********************
  '     PDesc225  ->  StkBySize
  '     PValue3   ->  StkByPtr
  '     PValue2   ->  DsgPtr
  '
  'qRngSz:  Join with RmRt table only if StkBySize = 'Y' or StkByPtr = 'Y'
  '**************** Logic For Getting qRngSz fields ********************

  '*** (The qSzYN in the SYyMm part of the query is special)
  
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDFR")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDFR")) <> "", " OdOmCmCd >= '" + Trim(adc("UOCMCDFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDTO")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDTO")) <> "", " OdOmCmCd <= '" + Trim(adc("UOCMCDTO")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UOCMCDSEL")) <> "", " And ", "") + IIF(Trim(adc("UOCMCDSEL")) <> "", " OdOmCmCd In (" + Trim(adc("UOCMCDSEL")) + ")", "")
  
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTFR")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTFR")) <> "01/01/80", " OdOmDt >= '" + Trim(adc("UODTFR")) + "'", "")
  wAddlCnd = wAddlCnd + IIF(wAddlCnd <> "" And Trim(adc("UODTTO")) <> "01/01/80", " And ", "") + IIF(Trim(adc("UODTTO")) <> "01/01/80", " OdOmDt <= '" + Trim(adc("UODTTO")) + "'", "")
  '****** Sachin 3.02.0 - Additional Conditions For New Fields
  
  '**** Sachin 2.12 - 02/12/05 - Silver added to the query for SJM
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc added
  ' ***** Manali 3.8.0 - Qty added
  ws_SqlStrg = "Select " + grpflds0 + ", " + _
               "SLoc, SRmCd as RmCd, max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg, " + _
               "max(case when SCTG.PValue3= 'Y' then convert(decimal(12, 4), SRmSz) else convert(decimal(12, 3), SRmSz) end) as qStkSz, " + _
               "max(case when SCTG.PValue3= 'Y' then convert(decimal(12, 4), SRmSz2) else convert(decimal(12, 3), SRmSz2) end) as qStkSz2, " + _
               "max(case when SCTG.PValue3= 'Y' then convert(decimal(12, 4), SRmSz3) else convert(decimal(12, 3), SRmSz3) end) as qStkSz3, " + _
               wRmSzDesc + " as qRmSzDesc, max(IsNull(RrRngSz, '')) as qRngSz, " + _
               "max(case when SCTG.PDesc225= 'Y' then SCTG.PDesc225 else (case when SCTG.PValue3= 'Y' then 'N' else 'Y' end) end) as qSzYN, " + _
               "sum(SRmDrWt- SRmCrWt) as qAvlblWt, sum(SRmDrQty- SRmCrQty) as qAvlblQty, 0  as qOrRmWt, 0 as qOrRmQty, 0 as qBalPrdQty, " + _
               "sum(SRmDrWt- SRmCrWt) * " + _
               "(Case when max(RmCtg) in ('G', 'P', 'S', 'L') then max(RmPurityWt) " + _
               "when max(RmCtg)= 'X' then max(RmBasePurityWt) else 0 end)/  " + _
               "(case max(RmCtg) when 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "When 'S' then " + CStr(wd_BaseSilRmPurity) + " " + _
               "When 'L' then " + CStr(wd_BasePdRmPurity) + " " + _
               "When 'X' then (case IsNull((Select RmCtg From RmMst " + _
               "               where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= max(Rm.RmBaseCd)), '') " + _
               "               When 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "               When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "               When 'S' then " + CStr(wd_BaseSilRmPurity) + " " + _
               "               When 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end) " + _
               "else 1 end) as qRmPureWt, 0 as qOrRmPureWt, " + _
               "'' as OrTc, '' as OrYy, '' as OrChr, '' as OrNo, 0 as OrSr, 0 as OrSrNo "
  ws_SqlStrg = ws_SqlStrg + " From SYyMm " + _
               "join Loc on LocCoCd= SCoCd and LocCd= SLoc " + _
               "join RmMst Rm on RmCd= SRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=SPrtKey ", "") + _
               "join Param SCTG on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg "
  
  Dim ws_Cnd1 As String, ws_Cnd2 As String
  
  ws_Cnd1 = GetCnd("R")
  ws_Cnd2 = GetCnd("D")

  ' ***** Manali 3.03 - 09/07/08 - join with RmIdSz added
  ws_SqlStrg = ws_SqlStrg + " left Outer join RmRt on RrTcTyp= 'RSZ' and RrCmCtg= 'C' " + _
               "and RrCmCd= 'ZSELF' and RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' " + _
               "and RrCd= '' and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' " + _
               "and RrFrLn<= (case when PDesc225 = 'Y' then convert (decimal(12, 3), SRmSz) when PValue3= 'Y' then convert (decimal(12, 4), SRmSz) else 999999 end) " + _
               "and RrToLn>= (case when PDesc225 = 'Y' then convert (decimal(12, 3), SRmSz) when PValue3= 'Y' then convert (decimal(12, 4), SRmSz) else 999999 end) " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=SRmSz " + _
               " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
               "where LocTyp= 'S' and SCoCd= '" + adc("UoCoCdFr") + "' " + _
               "and RmCtg= '" + adc("UoRmCtgFr") + "' " + ws_Cnd1
               
  '****** Sachin 2.12   21/12/05  - Show all RM or Order Specific
  '****** Sachin 3.02 - Id fields in Join
            If adc("UOYN1") = "N" Then ws_SqlStrg = ws_SqlStrg + " and exists((Select 'x' from Ordrm " + _
               "   Join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc  and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
               "   Join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc  and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
               "   Join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=DmPrtKey ", "") + _
               "   Join RmMst on RmCd=OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
               " where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + " OrCoCd=SCoCd And OrRmCd=SRmCd And OrLn1=SRmSz And OrLn2=SRmSz2 And OrLn3=SRmSz3 " + ws_Cnd2 + _
               IIF(wAddlCnd <> "", IIF(ws_Cnd2 <> "", " And ", " Where ") + wAddlCnd, "") + ")) "
  '****** Sachin 2.12   21/12/05  - Show all RM or Order Specific
  '' ***** Manali 3.8.0
  ws_SqlStrg = ws_SqlStrg + " Group By SLoc, SRmCd, convert(decimal(12, 4), SRmSz), convert(decimal(12, 4), SRmSz2), convert(decimal(12, 4), SRmSz3) " + _
               "having " + IIF(adc("UoYn3") = "N", " convert(decimal(16, 3), sum(SRmDrWt- SRmCrWt)) > 0  ", " sum(SRmDrQty- SRmCrQty) > 0 ")
               
' *** Bef 3.8.0
''  ws_SqlStrg = ws_SqlStrg + " Group By SLoc, SRmCd, convert(decimal(12, 4), SRmSz) " + _
''               "having convert(decimal(16, 3), sum(SRmDrWt- SRmCrWt)) > 0  "

               'Grp:  (case when PValue3= 'Y' then convert(decimal(12, 4), SRmSz) else convert(decimal(12, 3), SRmSz) end)

               '"and RrFrLn<= (case when (PDesc225 = 'Y' or PValue3= 'Y') then SRmSz else 999999 end) " + _
               "and RrToLn>= (case when (PDesc225 = 'Y' or PValue3= 'Y') then SRmSz else 999999 end) "


''  ws_RepCnd = adc.RepCond
''  ws_Cnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")

  ws_Cnd1 = GetCnd("D")

  '**************** Logic For Getting qStkSz & qRngSz fields ********************
  '     PDesc225  ->  StkBySize
  '     PValue3   ->  StkByPtr
  '     PValue2   ->  DsgPtr
  '
  'qStkSz:  If StkBySize = 'Y' then OrLn1
  '         Else
  '           If StkByPtr = 'Y' Then
  '             If OrPrdQty> 0 and OrRmCtg in('C','D') and HDTPrdPtr= 'Y' and DsgPtr= 'N' Then
  '               OrPrdWt/OrPrdQty
  '             Else
  '               OrRmPtr
  '             End If
  '           Else
  '             0
  '           End If
  '         End If
  '
  'Similarly for qRngSz
  '**************** Logic For Getting qStkSz & qRngSz fields ********************
  ' RamRam
  ' ***** Manali 3.03 - 09/07/08 - RmSzDesc added
  '6.1B-43 if exact by size option is selected then it will compare sz1,sz2,sz3 in txnd table to get balance qty
  'if range size is selected then it checks rm range size
  Dim wSzCnd As String, ws_RngSzJn As String
  If adc("UoYN2") = "N" Then
    wSzCnd = " and TDRMSZ=OrLn1 and TDRMSZ2 =OrLn2 and TDRMSZ3=OrLn3 "
    ws_RngSzJn = ""
  ElseIf adc("UoYN2") = "" Then
    ws_RngSzJn = " Left Outer Join RmRt ForBalQty On ForBalQty.RrTcTyp='RSZ' and ForBalQty.RrCmCtg= 'C' and " + _
                 "ForBalQty.RrCmCd='" + ctSelfCmCd + "' and ForBalQty.RrCtg=TdRmCtg and ForBalQty.RrSCtg=TdRmSCtg " + _
                 "and ForBalQty.RrSTWGrp= '' and ForBalQty.RrCd= '' and ForBalQty.RrDmCtg= '' and ForBalQty.RrLsCtg= '' " + _
                 "and ForBalQty.RrLabMCd= '' And ForBalQty.RrFrLn<= TdRmSz and ForBalQty.RrToLn>= TdRmSz "
    
    wSzCnd = " and isnull(ForBalQty.RrRngSz,'') = isnull(RmRt.RrRngSz,'') "
  End If
  '6.1B-43 if show 0 is NO then it hide the req wt/qty=0
  If adc("UoYN4") <> "Y" Then GRP_REP.SetFormula Rep, "wSwZeroReq", "'N'"
    
  '6.1B-43 wSzCnd,ws_RngSzJn added
  ws_SqlStrg = ws_SqlStrg + " Union All Select * From ( " + _
               "Select " + grpflds1 + ", " + _
               "'' as SLoc, OrRmCd as RmCd, OrRmCtg as RmCtg, OrRmSCtg as RmSCtg, " + _
               "(case P1.PDesc225 When 'Y' then convert (decimal(12, 3), OrLn1) " + _
               "else (case P1.PValue3 when 'Y' then " + _
               "        convert (decimal(12, 4), (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') " + _
               "              and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end)) else 0 end) " + _
               "end) as qStkSz, convert (decimal(12, 4), OrLn2) as qStkSz2, convert (decimal(12, 4), OrLn3) as qStkSz3, (Case when P1.PDesc225= 'Y' then IsNull(RisRmSzDesc, '') Else '' End) as qRmSzDesc, IsNull(RrRngSz, '') as qRngSz, " + _
               "(case when P1.PDesc225= 'Y' then P1.PDesc225 else (case when P1.PValue3= 'Y' then 'N' else 'Y' end) end) as qSzYN, " + _
               "0 as qAvlblWt, 0 as qAvlblQty, OrPrdWt as qOrRmWt, OrPrdQty as qOrRmQty, " + _
               "(OdPrdQty - isnull((Select sum(BQty) from Bag  where " + _
               IIF(gs_Partition = ctCurrPrtn, " BPrtKey=OrPrtKey and ", "") + " BCoCd = OrCoCd And BOdTc = OrTc And BOdYy = OrYy And BOdChr = OrChr " + _
               "and BOdNo= OrNo and BOdSr= OrSr " + _
               "and ((BCls= 'Y') or (BLoc= '" + ctFgLoc + "') " + _
               "or  ( convert(decimal(16, 3), isnull((select sum(TdRmWt* (Case " + _
               "      when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 " + _
               "      when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
               "      when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 " + _
               "      when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end)) " + _
               "      from Txnd " + ws_RngSzJn + " where TdPrtKey=BPrtKey and TdCoCd= BCoCd and TdBYy= BYy  " + _
               "      and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd " + wSzCnd + ") , 0.0)) > 0 " + _
               "    ) )), 0.0)) as qBalPrdQty, 0 as qRmPureWt, "
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  ws_SqlStrg = ws_SqlStrg + " OrPrdWt * (Case when RmCtg in ('G', 'P', 'S', 'L') then RmPurityWt " + _
               "      when RmCtg= 'X' then RmBasePurityWt else 0 end)/  " + _
               "      (case RmCtg when 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "              When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "              When 'S' then " + CStr(wd_BaseSilRmPurity) + " " + _
               "              When 'L' then " + CStr(wd_BasePdRmPurity) + " " + _
               "              When 'X' then (case IsNull((Select RmCtg From RmMst " + _
               "                             where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd= Rm.RmBaseCd), '') " + _
               "                             When 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
               "                             When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
               "                             When 'S' then " + CStr(wd_BaseSilRmPurity) + " " + _
               "                             When 'L' then " + CStr(wd_BasePdRmPurity) + " else 1 end " + _
               "      ) else 1 end) as qOrRmPureWt, " + _
               "OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo "
  '****** Sachin 3.02 - Id fields in Join
  ' ********* Manali 3.03 - Join with RmIdSz added
  '6.1-98 left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls added
  ws_SqlStrg = ws_SqlStrg + "From Param P join OrdRm on P.PTyp= 'TC' and P.PValue= 'SO' and OrTc= P.PMCd " + _
               "join OrdDsg on OdIdNo=OrOdIdNo and OdCoCd= OrCoCd and OdTc= OrTc  " + _
               "and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=OrPrtKey ", "") + _
               "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc  " + _
               "and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
               "join RmMst Rm on RmCd= OrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=OrPrtKey ", "") + _
               "join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
               "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
               "left outer join RmRt on  RrTcTyp= 'RSZ' and RrCmCtg= 'C' " + _
               "and RrCmCd= 'ZSELF' and RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' " + _
               "and RrCd= '' and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' " + _
               "and RrFrLn<= (case when P1.PDesc225 = 'Y' then convert (decimal(12, 3), OrLn1) else (case when P1.PValue3= 'Y' then convert (decimal(12, 4), (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end)) else 999999 end) end) " + _
               "and RrToLn>= (case when P1.PDesc225 = 'Y' then convert (decimal(12, 3), OrLn1) else (case when P1.PValue3= 'Y' then convert (decimal(12, 4), (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end)) else 999999 end) end) " + _
               "Left Outer Join RmIdSz on RisRmCtg=RmCtg and RisRmId=RmId and RisSz1=OrLn1 " + _
               " left outer join Param SC on SC.PTyp= 'STNCLS' and SC.PMCD=RmCtg and SC.PSCd=RmStnCls " + _
               "where " + IIF(gs_Partition = ctCurrPrtn, " OrPrtKey='" + ctCurrPrtn + "' and ", "") + " OrCoCd= '" + adc("UoCoCdFr") + "' and OrRmCtg= '" + adc("UoRmCtgFr") + "' " + _
               "and convert(decimal(12, 1), (OdPrdQty- OdFgQty))> 0 " + _
               "and " + IIF(adc("UoYn3") = "N", " convert(decimal(16, 3), OrPrdWt)> 0 ", " OrPrdQty > 0") + _
               ws_Cnd1 + IIF(wAddlCnd <> "", IIF(ws_Cnd1 <> "", " And ", " Where ") + wAddlCnd, "") + ") RmReq "
               
               '''' Bef 3.8.0
               '''"and convert(decimal(16, 3), OrPrdWt)> 0 " + ws_Cnd1 + IIF(wAddlCnd <> "", IIF(ws_Cnd1 <> "", " And ", " Where ") + wAddlCnd, "") + ") RmReq "
  ws_SqlStrg = ws_SqlStrg + " where  convert(decimal(12, 1), qBalPrdQty) <> 0 " + _
               "Order By RmCd, qStkSz, qStkSz2, qStkSz3, OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo, SLoc " + _
               "OPTION (FORCE ORDER, MAXDOP 1) "
               '****** Sachin 3.01 added Maxdop Option above

              '"and RrFrLn<= (case when P1.PDesc225 = 'Y' then OrLn1 else (case when P1.PValue3= 'Y' then (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end) else 999999 end) end) " + _
               "and RrToLn>= (case when P1.PDesc225 = 'Y' then OrLn1 else (case when P1.PValue3= 'Y' then (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end) else 999999 end) end) "
    
  adc.RepRecSource = ws_SqlStrg
  
  ' ******* Manali 3.03 - 09/07/08 - RmSzDesc
'  GRP_REP.SetFormula Rep, "wStkSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} " + _
                                    "Else (If {rdo.qSzYN}= 'Y' then ToText({rdo.qStkSz}, 3) Else ToText({rdo.qStkSz}, 4))"
  GRP_REP.SetFormula Rep, "wStkSz", "If {rdo.qRmSzDesc}<> '' Then {rdo.qRmSzDesc} Else ''"
                                    
   ' ******* Manali 3.03 - 09/07/08 - RmSzDesc
  
  
  ' ***** Manali 3.8.0
  Dim ws_QW As String
  
  If adc("UoYn3") = "Y" Then
    ws_QW = "Qty"
    GRP_REP.SetFormula Rep, "wAvlblPureWt", "0":    GRP_REP.SetFormula Rep, "wBalReqPureWt", "0"
    GRP_REP.SetFormula Rep, "g3_AvlblPureWt", "0"
    GRP_REP.SetFormula Rep, "g2_AvlblPureWt", "0"
    GRP_REP.SetFormula Rep, "g1_AvlblPureWt", "0"
    GRP_REP.SetFormula Rep, "g_AvlblPureWt", "0"
    
    Rep.Fldg4AvlblWt.DecimalPlaces = 0: Rep.Fldg4AvlblWt.RoundingType = crRoundToUnit
    Rep.Fldg3AvlblWt.DecimalPlaces = 0: Rep.Fldg3AvlblWt.RoundingType = crRoundToUnit
    Rep.Fldg2AvlblWt.DecimalPlaces = 0: Rep.Fldg2AvlblWt.RoundingType = crRoundToUnit
    Rep.Fldg1AvlblWt.DecimalPlaces = 0: Rep.Fldg1AvlblWt.RoundingType = crRoundToUnit
    Rep.FldgAvlblWt.DecimalPlaces = 0: Rep.FldgAvlblWt.RoundingType = crRoundToUnit

    Rep.Fldg4wReqWt.DecimalPlaces = 0: Rep.Fldg4wReqWt.RoundingType = crRoundToUnit
    Rep.Fldg3wReqWt.DecimalPlaces = 0: Rep.Fldg3wReqWt.RoundingType = crRoundToUnit
    Rep.Fldg2wReqWt.DecimalPlaces = 0: Rep.Fldg2wReqWt.RoundingType = crRoundToUnit
    Rep.Fldg1wReqWt.DecimalPlaces = 0: Rep.Fldg1wReqWt.RoundingType = crRoundToUnit
    Rep.FldgwReqWt.DecimalPlaces = 0: Rep.FldgwReqWt.RoundingType = crRoundToUnit

    Rep.Fldg4NetWt.DecimalPlaces = 0: Rep.Fldg4NetWt.RoundingType = crRoundToUnit
    Rep.Fldg3NetWt.DecimalPlaces = 0: Rep.Fldg3NetWt.RoundingType = crRoundToUnit
    Rep.Fldg2NetWt.DecimalPlaces = 0: Rep.Fldg2NetWt.RoundingType = crRoundToUnit
    Rep.Fldg1NetWt.DecimalPlaces = 0: Rep.Fldg1NetWt.RoundingType = crRoundToUnit
    Rep.FldgNetWt.DecimalPlaces = 0: Rep.FldgNetWt.RoundingType = crRoundToUnit

    Rep.FldBalReqWt.DecimalPlaces = 0: Rep.FldBalReqWt.RoundingType = crRoundToUnit
    Rep.FldOrRmWt.DecimalPlaces = 0: Rep.FldOrRmWt.RoundingType = crRoundToUnit
    Rep.FldDetAvlblWt.DecimalPlaces = 0: Rep.FldDetAvlblWt.RoundingType = crRoundToUnit
    
    Rep.hAvlblWt.SetText "Avlbl Qty": Rep.h1AvlblWt.SetText "Avlbl Qty"
    Rep.hReqWt.SetText "Req Qty": Rep.h1ReqWt.SetText "Req Qty"
    Rep.hNetWt.SetText "Net Qty": Rep.h1NetWt.SetText "Net Qty"
    Rep.LblAvlblWt.SetText "Avlbl Qty": Rep.LblReqWt.SetText "Req Qty": Rep.LblWtperPc.SetText "Qty/ Pc"
    Rep.hAvlblPureWt.Suppress = True: Rep.h1AvlblPureWt.Suppress = True
    Rep.hReqPureWt.Suppress = True: Rep.h1ReqPureWt.Suppress = True
    Rep.hNetPureWt.Suppress = True: Rep.h1NetPureWt.Suppress = True
    
  Else
    ws_QW = "Wt"
   
    Rep.hAvlblWt.SetText "Avlbl Wt": Rep.h1AvlblWt.SetText "Avlbl Wt"
    Rep.hReqWt.SetText "Req Wt": Rep.h1ReqWt.SetText "Req Wt"
    Rep.hNetWt.SetText "Net Wt": Rep.h1NetWt.SetText "Net Wt"
    Rep.LblAvlblWt.SetText "Avlbl Wt": Rep.LblReqWt.SetText "Req Wt": Rep.LblWtperPc.SetText "Wt/ Pc"
    Rep.hAvlblPureWt.Suppress = False: Rep.h1AvlblPureWt.Suppress = False
    Rep.hReqPureWt.Suppress = False: Rep.h1ReqPureWt.Suppress = False
    Rep.hNetPureWt.Suppress = False: Rep.h1NetPureWt.Suppress = False
  End If
  '6.1B-43
  Dim wFormulaStr As String
  If adc("UoYN2") = "N" Then
    wFormulaStr = "{rdo.RmCd} + ToText ({rdo.qStkSz}, 4)+ ToText ({rdo.qStkSz2}, 4)+ ToText ({rdo.qStkSz3}, 4)"
    GRP_REP.SetFormula Rep, "wGrp4", wFormulaStr
  End If
 
  GRP_REP.SetFormula Rep, "wAvlblWt", "Sum ({rdo.qAvlbl" + ws_QW + "}, {@wGrp4})"
  GRP_REP.SetFormula Rep, "wBalReqWt", "{rdo.qOrRm" + ws_QW + "}* {rdo.qBalPrdQty}"
  GRP_REP.SetFormula Rep, "wOrRmWt", "{rdo.qOrRm" + ws_QW + "}"
  GRP_REP.SetFormula Rep, "wDetAvlblWt", "{rdo.qAvlbl" + ws_QW + "}"
  
  GRP_REP.SetFormula Rep, "g3_AvlblWt", "Sum ({rdo.qAvlbl" + ws_QW + "}, {rdo.Grp3})"
  GRP_REP.SetFormula Rep, "g2_AvlblWt", "Sum ({rdo.qAvlbl" + ws_QW + "}, {rdo.Grp2})"
  GRP_REP.SetFormula Rep, "g1_AvlblWt", "Sum ({rdo.qAvlbl" + ws_QW + "}, {rdo.Grp1})"
  GRP_REP.SetFormula Rep, "g_AvlblWt", "Sum ({rdo.qAvlbl" + ws_QW + "})"
  
  ' ***** Manali 3.8.0
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 8)
  
  Dim ws_LocScope As String
  ws_LocScope = IIF(adc("UoIssLocFr") = "", "", "From: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocTo") = "", "", "   To: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocSel") = "", "", "    Selection: " + adc("UoIssLocSel"))
  If ws_LocScope <> "" Then ws_LocScope = "Loc " + ws_LocScope
  Rep.TxtLocScope.SetText ws_LocScope
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

Private Sub adc_setreprecsource_1()
'  '*** Report Sql ***
'  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  '  SLoc, SRmCd as RmCd, OrRmCtg as RmCtg, OrRmSCtg as RmSCtg,
'  '  SRmSz as qStkSz, RrRngSz as qRngSz, space(1) as qSzYN,
'  '  SRmDrWt as qAvlblWt, SRmDrWt as qOrRmWt, OdPrdQty as qBalPrdQty,
'  '  SRmDrWt As qRmPureWt, SRmDrWt As qOrRmPureWt,
'  '  OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo
'  '  From SYyMm, OrdRm, OrdDsg, RmRt
'  '  Where 1 = 2
'  '*** Report Sql ***
'
'
'  Dim grpflds0 As String, grpflds1 As String
'  Dim ws_RepCnd As String, ws_Cnd As String, ws_SqlStrg As String
'
'  Dim wd_BaseGldRmPurity As Double, wd_BasePltRmPurity As Double
'
'  Dim ws_HDTPrdPtr As String
'
'  Set REP = Nothing
'  Set adc.RepSource = REP
'  Call DispCoNm
'  Call GRP_REP.Gen3LRep
'
'  grpflds0 = GRP_REP.GrpFldLst(0)
'  grpflds1 = GRP_REP.GrpFldLst(1)
'
'
'  '*** To check if Detail is selected ***
'  Dim ws_Grp(6) As String, wi_i As Integer, wi_DetPos As Integer
'  ws_Grp(1) = GRP_REP.Value(gltGroup1)
'  ws_Grp(2) = GRP_REP.Value(gltGroup2)
'  ws_Grp(3) = GRP_REP.Value(gltGroup3)
'  wi_DetPos = 0
'  For wi_i = 1 To 6
'    If ws_Grp(wi_i) = "(Detail)" Then
'      wi_DetPos = wi_i: Exit For
'    End If
'  Next wi_i
'
'  If wi_DetPos = 0 Or adc("UoYn") = "N" Then
'    REP.BoxHdLoc.Suppress = True: REP.BoxDetLoc.Suppress = True: REP.BoxFtLoc.Suppress = True
'    REP.BoxHdOrd.Suppress = True: REP.BoxDetOrd.Suppress = True: REP.BoxFtOrd.Suppress = True
'    If wi_DetPos = 0 Then REP.G4F.Suppress = True
'  End If
'
'  wd_BaseGldRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'G'), 1) ")
'  wd_BasePltRmPurity = moCn.GetFldVal("Select IsNull((select RmPurityWt from Param join RmMst on RmCd= PValue where PTyp= 'RMCTG' and PMCd= 'P'), 1) ")
'  ws_HDTPrdPtr = UCase(moCn.GetFldVal("Select HDTPrdPtr from Head where HCoCd= '" + adc("UoCoCdFr") + "'"))
'
'  'adc("UoLn1Fr").CmpStr = "SRmSz>= ": adc("UoLn1To").CmpStr = "SRmSz<= "
'  adc("UoOmTcFr").CmpStr = "": adc("UoOmTcTo").CmpStr = ""
'  adc("UoOmYyFr").CmpStr = "": adc("UoOmYyTo").CmpStr = ""
'  adc("UoOmChrFr").CmpStr = "": adc("UoOmChrTo").CmpStr = ""
'  adc("UoOmNoFr").CmpStr = "": adc("UoOmNoTo").CmpStr = ""
'  adc("UoOdSrFr").CmpStr = "": adc("UoOdSrTo").CmpStr = ""
'  adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
'  adc("UoDmCdFr").CmpStr = "": adc("UoDmCdTo").CmpStr = "": adc("UoDmCdSel").CmpStr = ""
'  adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
'  adc("UoPoNoFr").CmpStr = "": adc("UoPoNoTo").CmpStr = "": adc("UoPoNoSel").CmpStr = ""
'  adc("UoDtFr").CmpStr = "": adc("UoDtTo").CmpStr = ""
'  adc("UoDelDtFr").CmpStr = "": adc("UoDelDtTo").CmpStr = ""
'  adc("UoPrtCdFr").CmpStr = "": adc("UoPrtCdTo").CmpStr = "": adc("UoPrtCdSel").CmpStr = ""
'  adc("UoIssLocFr").CmpStr = "SLoc >=": adc("UoIssLocTo").CmpStr = "SLoc <=": adc("UoIssLocSel").CmpStr = "SLoc In"
'
'  ws_RepCnd = adc.RepCond
'  ws_Cnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
'
'  '**************** Logic For Getting qRngSz fields ********************
'  '     PDesc225  ->  StkBySize
'  '     PValue3   ->  StkByPtr
'  '     PValue2   ->  DsgPtr
'  '
'  'qRngSz:  Join with RmRt table only if StkBySize = 'Y' or StkByPtr = 'Y'
'  '**************** Logic For Getting qRngSz fields ********************
'
'  '*** (The qSzYN in the SYyMm part of the query is special)
'  ws_SqlStrg = "Select " + grpflds0 + ", " + _
'               "SLoc, SRmCd as RmCd, max(RmCtg) as RmCtg, max(RmSCtg) as RmSCtg, " + _
'               "SRmSz as qStkSz, max(IsNull(RrRngSz, '')) as qRngSz, " + _
'               "max(case when PDesc225= 'Y' then PDesc225 else (case when PValue3= 'Y' then 'N' else 'Y' end) end) as qSzYN, " + _
'               "sum(SRmDrWt- SRmCrWt) as qAvlblWt, 0  as qOrRmWt, 0 as qBalPrdQty, " + _
'               "sum(SRmDrWt- SRmCrWt) * " + _
'               "(Case when max(RmCtg) in ('G', 'P') then max(RmPurityWt) " + _
'               "when max(RmCtg)= 'X' then max(RmBasePurityWt) else 0 end)/  " + _
'               "(case max(RmCtg) when 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
'               "When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
'               "When 'X' then (case IsNull((Select RmCtg From RmMst " + _
'               "               where RmCd= max(Rm.RmBaseCd)), '') " + _
'               "               When 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
'               "               When 'P' then " + CStr(wd_BasePltRmPurity) + " else 1 end) " + _
'               "else 1 end) as qRmPureWt, 0 as qOrRmPureWt, " + _
'               "'' as OrTc, '' as OrYy, '' as OrChr, '' as OrNo, 0 as OrSr, 0 as OrSrNo "
'  ws_SqlStrg = ws_SqlStrg + " From SYyMm " + _
'               "join Loc with (index(PK_Loc)) on LocCoCd= SCoCd and LocCd= SLoc " + _
'               "join RmMst Rm on RmCd= SRmCd " + _
'               "join Param on PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg " + _
'               "left Outer join RmRt on RrTcTyp= 'RSZ' and RrCmCtg= 'C' " + _
'               "and RrCmCd= 'ZSELF' and RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' " + _
'               "and RrCd= '' and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' " + _
'               "and RrFrLn<= (case when (PDesc225 = 'Y' or PValue3= 'Y') then SRmSz else 999999 end) " + _
'               "and RrToLn>= (case when (PDesc225 = 'Y' or PValue3= 'Y') then SRmSz else 999999 end) " + _
'               "where LocTyp= 'S' and SCoCd= '" + adc("UoCoCdFr") + "' " + _
'               "and RmCtg= '" + adc("UoRmCtgFr") + "' " + ws_Cnd + _
'               "Group By SLoc, SRmCd, SRmSz " + _
'               "having convert(decimal(16, 3), sum(SRmDrWt- SRmCrWt)) > 0  "
'
'
'  'adc("UoLn1Fr").CmpStr = "OrLn1>= ": adc("UoLn1To").CmpStr = "OrLn1<= "
'  adc("UoOmTcFr").CmpStr = "OrTc>= ": adc("UoOmTcTo").CmpStr = "OrTc<= "
'  adc("UoOmYyFr").CmpStr = "OrYy>= ": adc("UoOmYyTo").CmpStr = "OrYy<= "
'  adc("UoOmChrFr").CmpStr = "OrChr>= ": adc("UoOmChrTo").CmpStr = "OrChr<= "
'  adc("UoOmNoFr").CmpStr = "OrNo>= ": adc("UoOmNoTo").CmpStr = "OrNo<= "
'  adc("UoOdSrFr").CmpStr = "OrSr>= ": adc("UoOdSrTo").CmpStr = "OrSr<= "
'  adc("UoDmCtgFr").CmpStr = "DmCtg >=": adc("UoDmCtgTo").CmpStr = "DmCtg <=": adc("UoDmCtgSel").CmpStr = "DmCtg In"
'  adc("UoDmCdFr").CmpStr = "OdDmCd >=": adc("UoDmCdTo").CmpStr = "OdDmCd <=": adc("UoDmCdSel").CmpStr = "OdDmCd In"
'  adc("UoCmCdFr").CmpStr = "OmCmCd >=": adc("UoCmCdTo").CmpStr = "OmCmCd <=": adc("UoCmCdSel").CmpStr = "OmCmCd In"
'  adc("UoPoNoFr").CmpStr = "OmPoNo >=": adc("UoPoNoTo").CmpStr = "OmPoNo <=": adc("UoPoNoSel").CmpStr = "OmPoNo In"
'  adc("UoDtFr").CmpStr = "OmDt >=": adc("UoDtTo").CmpStr = "OmDt <="
'  adc("UoDelDtFr").CmpStr = "OdDelDt >=": adc("UoDelDtTo").CmpStr = "OdDelDt <="
'  adc("UoPrtCdFr").CmpStr = "OdPrtCd >=": adc("UoPrtCdTo").CmpStr = "OdPrtCd <=": adc("UoPrtCdSel").CmpStr = "OdPrtCd In"
'  adc("UoIssLocFr").CmpStr = "": adc("UoIssLocTo").CmpStr = "": adc("UoIssLocSel").CmpStr = ""
'
'  ws_RepCnd = adc.RepCond
'  ws_Cnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
'
'  '**************** Logic For Getting qStkSz & qRngSz fields ********************
'  '     PDesc225  ->  StkBySize
'  '     PValue3   ->  StkByPtr
'  '     PValue2   ->  DsgPtr
'  '
'  'qStkSz:  If StkBySize = 'Y' then OrLn1
'  '         Else
'  '           If StkByPtr = 'Y' Then
'  '             If OrPrdQty> 0 and OrRmCtg in('C','D') and HDTPrdPtr= 'Y' and DsgPtr= 'N' Then
'  '               OrPrdWt/OrPrdQty
'  '             Else
'  '               OrRmPtr
'  '             End If
'  '           Else
'  '             0
'  '           End If
'  '         End If
'  '
'  'Similarly for qRngSz
'  '**************** Logic For Getting qStkSz & qRngSz fields ********************
'
'  ws_SqlStrg = ws_SqlStrg + " Union All Select * From ( " + _
'               "Select " + grpflds1 + ", " + _
'               "'' as SLoc , OrRmCd as RmCd, OrRmCtg as RmCtg, OrRmSCtg as RmSCtg, " + _
'               "(case P1.PDesc225 When 'Y' then OrLn1 " + _
'               "else (case P1.PValue3 when 'Y' then " + _
'               "        (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') " + _
'               "              and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end) else 0 end) " + _
'               "end) as qStkSz, IsNull(RrRngSz, '') as qRngSz, " + _
'               "(case when P1.PDesc225= 'Y' then P1.PDesc225 else (case when P1.PValue3= 'Y' then 'N' else 'Y' end) end) as qSzYN, " + _
'               "0 as qAvlblWt, OrPrdWt as qOrRmWt, " + _
'               "(OdPrdQty - isnull((Select sum(BQty) from Bag with (index(xBOrdNo)) where " + _
'               "BCoCd = OrCoCd And BOdTc = OrTc And BOdYy = OrYy And BOdChr = OrChr " + _
'               "and BOdNo= OrNo and BOdSr= OrSr " + _
'               "and ((BCls= 'Y') or (BLoc= '" + ctFgLoc + "') " + _
'               "or  ( convert(decimal(16, 3), isnull((select sum(TdRmWt* (Case " + _
'               "      when TdFrRmLoc= 'W' and TdFrRmDc= 'D' then 1 " + _
'               "      when TdFrRmLoc= 'W' and TdFrRmDc= 'C' then -1 " + _
'               "      when TdToRmLoc= 'W' and TdToRmDc= 'D' then 1 " + _
'               "      when TdToRmLoc= 'W' and TdToRmDc= 'C' then -1 Else 0 end)) " + _
'               "      from Txnd with (index(xTdBag)) where TdCoCd= BCoCd and TdBYy= BYy  " + _
'               "      and TdBChr= BChr and TdBNo= BNo and TdRmCd= OrRmCd) , 0.0)) > 0 " + _
'               "    ) )), 0.0)) as qBalPrdQty, 0 as qRmPureWt, "
'  ws_SqlStrg = ws_SqlStrg + " OrPrdWt * (Case when RmCtg in ('G', 'P') then RmPurityWt " + _
'               "      when RmCtg= 'X' then RmBasePurityWt else 0 end)/  " + _
'               "      (case RmCtg when 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
'               "              When 'P' then " + CStr(wd_BasePltRmPurity) + " " + _
'               "              When 'X' then (case IsNull((Select RmCtg From RmMst " + _
'               "                             where RmCd= Rm.RmBaseCd), '') " + _
'               "                             When 'G' then " + CStr(wd_BaseGldRmPurity) + " " + _
'               "                             When 'P' then " + CStr(wd_BasePltRmPurity) + " else 1 end " + _
'               "      ) else 1 end) as qOrRmPureWt, " + _
'               "OrTc , OrYy, OrChr, OrNo, OrSr, OrSrNo "
'  ws_SqlStrg = ws_SqlStrg + "From Param P join OrdRm on P.PTyp= 'TC' and P.PValue= 'SO' and OrTc= P.PMCd " + _
'               "join OrdDsg with (index(PK_OrdDsg)) on OdCoCd= OrCoCd and OdTc= OrTc  " + _
'               "and OdYy= OrYy and OdChr= OrChr and OdNo= OrNo and OdSr= OrSr " + _
'               "join OrdMst with (index(PK_OrdMst)) on OmCoCd= OdCoCd and OmTc= OdTc  " + _
'               "and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo " + _
'               "join RmMst Rm on RmCd= OrRmCd " + _
'               "join Param P1 on P1.PTyp= 'RMSCTG' and P1.PMCd= RmCtg and P1.PSCd= RmSCtg " + _
'               "join DsgMst on DmTcTyp= 'DM' and DmCd= OdDmCd and DmSz= '' " + _
'               "left outer join RmRt on RrTcTyp= 'RSZ' and RrCmCtg= 'C' " + _
'               "and RrCmCd= 'ZSELF' and RrCtg= RmCtg and RrSCtg= RmSCtg and RrSTWGrp= '' " + _
'               "and RrCd= '' and RrDmCtg= '' and RrLsCtg= '' and RrLabMCd= '' " + _
'               "and RrFrLn<= (case when P1.PDesc225 = 'Y' then OrLn1 else (case when P1.PValue3= 'Y' then (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end) else 999999 end) end) " + _
'               "and RrToLn>= (case when P1.PDesc225 = 'Y' then OrLn1 else (case when P1.PValue3= 'Y' then (case when OrPrdQty> 0 and OrRmCtg in ('C', 'D') and '" + ws_HDTPrdPtr + "' = 'Y' and P1.PValue2= 'N' then OrPrdWt/OrPrdQty else OrRmPtr end) else 999999 end) end) " + _
'               "where OrCoCd= '" + adc("UoCoCdFr") + "' and OrRmCtg= '" + adc("UoRmCtgFr") + "' " + _
'               "and convert(decimal(12, 1), (OdPrdQty- OdFgQty))> 0 " + _
'               "and convert(decimal(16, 3), OrPrdWt)> 0 " + ws_Cnd + ") RmReq "
'  ws_SqlStrg = ws_SqlStrg + " where convert(decimal(12, 1), qBalPrdQty) <> 0 " + _
'               "Order By OrTc, OrYy, OrChr, OrNo, OrSr, OrSrNo " + _
'               "OPTION (FORCE ORDER) "
'
'
'  adc.RepRecSource = ws_SqlStrg
'
'  Dim ws_LocScope As String
'  ws_LocScope = IIF(adc("UoIssLocFr") = "", "", "From: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocTo") = "", "", "   To: " + adc("UoIssLocFr")) + IIF(adc("UoIssLocSel") = "", "", "    Selection: " + adc("UoIssLocSel"))
'  If ws_LocScope <> "" Then ws_LocScope = "Loc " + ws_LocScope
'  REP.TxtLocScope.SetText ws_LocScope
'
'  CRV_REP.DisplayGroupTree = False
End Sub

Private Function GetCnd(ByVal ws_Typ As String) As String
Dim ws_RepCnd As String
If ws_Typ = "R" Then
    adc("UoOmTcFr").CmpStr = "": adc("UoOmTcTo").CmpStr = ""
    adc("UoOmYyFr").CmpStr = "": adc("UoOmYyTo").CmpStr = ""
    adc("UoOmChrFr").CmpStr = "": adc("UoOmChrTo").CmpStr = ""
    adc("UoOmNoFr").CmpStr = "": adc("UoOmNoTo").CmpStr = "": adc("UoOmNoSel").CmpStr = ""   '578-28 OmNo Selection added
    adc("UoOdSrFr").CmpStr = "": adc("UoOdSrTo").CmpStr = ""
    adc("UoDmCtgFr").CmpStr = "": adc("UoDmCtgTo").CmpStr = "": adc("UoDmCtgSel").CmpStr = ""
    adc("UoDmCdFr").CmpStr = "": adc("UoDmCdTo").CmpStr = "": adc("UoDmCdSel").CmpStr = ""
    adc("UoCmCdFr").CmpStr = "": adc("UoCmCdTo").CmpStr = "": adc("UoCmCdSel").CmpStr = ""
    adc("UoPoNoFr").CmpStr = "": adc("UoPoNoTo").CmpStr = "": adc("UoPoNoSel").CmpStr = ""
    adc("UoOdSubCustFr").CmpStr = "": adc("UoOdSubCustTo").CmpStr = ""  'VK.19
    adc("UoOdPoNoFr").CmpStr = "": adc("UoOdPoNoTo").CmpStr = "": adc("UoOdPoNoSel").CmpStr = ""    '*** (Jen 3.2.0)
    adc("UoDtFr").CmpStr = "": adc("UoDtTo").CmpStr = ""
    adc("UoDelDtFr").CmpStr = "": adc("UoDelDtTo").CmpStr = ""
    adc("UoPrtCdFr").CmpStr = "": adc("UoPrtCdTo").CmpStr = "": adc("UoPrtCdSel").CmpStr = ""
    adc("UoIssLocFr").CmpStr = "SLoc >=": adc("UoIssLocTo").CmpStr = "SLoc <=": adc("UoIssLocSel").CmpStr = "SLoc In"
    ws_RepCnd = adc.RepCond
    GetCnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
    
ElseIf ws_Typ = "D" Then
    adc("UoOmTcFr").CmpStr = "OrTc>= ": adc("UoOmTcTo").CmpStr = "OrTc<= "
    adc("UoOmYyFr").CmpStr = "OrYy>= ": adc("UoOmYyTo").CmpStr = "OrYy<= "
    adc("UoOmChrFr").CmpStr = "OrChr>= ": adc("UoOmChrTo").CmpStr = "OrChr<= "
    adc("UoOmNoFr").CmpStr = "OrNo>= ": adc("UoOmNoTo").CmpStr = "OrNo<= "
    adc("UoOmNoSel").CmpStr = "OrNo In "       '578-28 OmNo Selection added
    adc("UoOdSrFr").CmpStr = "OrSr>= ": adc("UoOdSrTo").CmpStr = "OrSr<= "
    adc("UoDmCtgFr").CmpStr = "DmCtg >=": adc("UoDmCtgTo").CmpStr = "DmCtg <=": adc("UoDmCtgSel").CmpStr = "DmCtg In"
    adc("UoDmCdFr").CmpStr = "OdDmCd >=": adc("UoDmCdTo").CmpStr = "OdDmCd <=": adc("UoDmCdSel").CmpStr = "OdDmCd In"
    adc("UoCmCdFr").CmpStr = "OmCmCd >=": adc("UoCmCdTo").CmpStr = "OmCmCd <=": adc("UoCmCdSel").CmpStr = "OmCmCd In"
    adc("UoOdSubCustFr").CmpStr = "OdSubCust >=": adc("UoOdSubCustTo").CmpStr = "OdSubCust <="  'VK.19
    adc("UoPoNoFr").CmpStr = "OmPoNo >=": adc("UoPoNoTo").CmpStr = "OmPoNo <=": adc("UoPoNoSel").CmpStr = "OmPoNo In"
    adc("UoOdPoNoFr").CmpStr = "OdPoNo >=": adc("UoOdPoNoTo").CmpStr = "OdPoNo <=": adc("UoOdPoNoSel").CmpStr = "OdPoNo in"   '*** (Jen 3.2.0)
    adc("UoDtFr").CmpStr = "OmDt >=": adc("UoDtTo").CmpStr = "OmDt <="
    adc("UoDelDtFr").CmpStr = "OdDelDt >=": adc("UoDelDtTo").CmpStr = "OdDelDt <="
    adc("UoPrtCdFr").CmpStr = "OdPrtCd >=": adc("UoPrtCdTo").CmpStr = "OdPrtCd <=": adc("UoPrtCdSel").CmpStr = "OdPrtCd In"
    adc("UoIssLocFr").CmpStr = "": adc("UoIssLocTo").CmpStr = "": adc("UoIssLocSel").CmpStr = ""
    ws_RepCnd = adc.RepCond
    GetCnd = IIF(ws_RepCnd <> "", " and " + ws_RepCnd, "")
  End If

End Function


' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0       ' Tag = "ShowDet"
    If ChkBoxArr(0).Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1
    If ChkBoxArr(1).Value = Checked Then adc("UoYN4") = "Y" Else adc("UoYN4") = "N"   '6.1B-43
  End Select
End With
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' **** Manali 3.5.0 - 18/11/08 - Check Boxes Added

'*** std code not to be changed
Private Sub Form_GotFocus()
  '*** Jenny Color (For Setting System Colors)
  '(Bef 13/08/05)
  'SetMDISysColorsAfterLogin
  '*** Jenny Color (For Setting System Colors)
End Sub
Private Sub Form_Load()
  '*** Jenny Colour
  SetMDISysColorsAfterLogin   '*** Jenny Color (For Setting System Colors)
  
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
    gs_LocTyp = "'S'"
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '***  Extra Validation Code Here
  '578-28 validating order no. selection should be a number
  Select Case UCase(IdName)
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
  '***  Extra Field When Code Here
  gb_CoCdFor = True         '*** (Jen 2.14 Next Patch 1)
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
  '*** Recalc Code Here
End Sub

' **** Manali 3.8.0 - Qty Or Wt
Private Sub OptQW_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN3") = "Y"
Case Is = 1
  adc("UoYN3") = "N"
End Select
End Sub
' **** Manali 3.8.0 - Qty Or Wt

' **** Manali 3.8.0 - Qty Or Wt
Private Sub OptQW_GotFocus(Index As Integer)
  DispMsg FraOptQW.ToolTipText, etInfo
End Sub

' **** Manali 3.5.0 - 19/11/08 - Radio Buttons Added
Private Sub OptRm_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
End Select
End Sub
' **** Manali 3.5.0 - 19/11/08 - Radio Buttons Added

' **** Manali 3.5.0 - 19/11/08 - Radio Buttons Added
Private Sub OptRm_GotFocus(Index As Integer)
  DispMsg FraOptRm.ToolTipText, etInfo
End Sub
' **** Manali 3.5.0 - 19/11/08 - Radio Buttons Added

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
  ' Zubin 212
  If adc("UOYN1") = "" Then adc("UOYN1") = "Y"
  ' Zubin 212
  
  ' **** Manali 3.5.0 - 19/11/08 - radio Buttons and Check Boxes Added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then OptRm(0).Value = True
  If adc("UoYN1") = "N" Then OptRm(1).Value = True
  ' **** Manali 3.5.0 - 19/11/08 - Radio buttons and Check Boxes Added
  
  If adc("UoYN3") = "Y" Then OptQW(0).Value = True
  If adc("UoYN3") = "N" Then OptQW(1).Value = True
  '6.1B-43 default value for the report is Rmcode
  adc("UoYN2") = "Y"
  OptRpt(0).Value = True
  '6.1B-43 if show 0 option is blank then it set as No
  If adc("UOYN4") = "" Then adc("UOYN4") = "N"
  If adc("UoYN4") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
    
End Sub
Private Sub DispCoNm()
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
'6.1B-43 report options selection
Private Sub OptRpt_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    adc("UoYN2") = "Y"
  Case Is = 1
    adc("UoYN2") = "N"
  Case Is = 2
    adc("UoYN2") = ""
  End Select
End Sub
'6.1B-43 tooltip for report options
Private Sub OptRpt_GotFocus(Index As Integer)
  DispMsg OptRpt(Index).ToolTipText, etInfo
End Sub

