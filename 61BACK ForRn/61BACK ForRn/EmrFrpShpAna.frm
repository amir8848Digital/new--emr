VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpShpAna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipment Analysis"
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
   Begin VB.Frame FRA_REPOBJ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      TabIndex        =   67
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   68
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
      Height          =   9585
      Left            =   0
      TabIndex        =   37
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   36
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpShpAna.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group&Sort"
         TabPicture(1)   =   "EmrFrpShpAna.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9030
            Index           =   1
            Left            =   -74940
            TabIndex        =   39
            Top             =   360
            Width           =   15090
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   9435
               _ExtentX        =   16642
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9045
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   15075
            Begin VB.Frame FraOptQty 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1965
               TabIndex        =   56
               ToolTipText     =   "Select Shipment Status All/Closed"
               Top             =   3630
               Width           =   4455
               Begin VB.OptionButton OptShp 
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
                  TabIndex        =   33
                  ToolTipText     =   "Select Shipment Status All/Closed"
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1335
               End
               Begin VB.OptionButton OptShp 
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
                  Left            =   2685
                  TabIndex        =   34
                  ToolTipText     =   "Select Shipment Status All/Closed"
                  Top             =   0
                  Width           =   1245
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   4680
               TabIndex        =   19
               ToolTipText     =   "Enter To Design Category"
               Top             =   2085
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
               Left            =   7410
               TabIndex        =   20
               ToolTipText     =   "Enter Design Category Selection"
               Top             =   2085
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
               Index           =   30
               Left            =   1965
               TabIndex        =   18
               ToolTipText     =   "Enter From Design Category"
               Top             =   2085
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
               Left            =   4680
               TabIndex        =   22
               ToolTipText     =   "Enter To Design Code"
               Top             =   2370
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
               Left            =   7410
               TabIndex        =   23
               ToolTipText     =   "Enter Design Code Selection"
               Top             =   2370
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UODMCDSEL"
               IdName          =   "UODMCDSEL"
               CmpStr          =   "OdDmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   33
               Left            =   1965
               TabIndex        =   21
               ToolTipText     =   "Enter From Design Code"
               Top             =   2370
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
               Index           =   40
               Left            =   4680
               TabIndex        =   16
               ToolTipText     =   "Enter To Customer"
               Top             =   1800
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
               Left            =   7410
               TabIndex        =   17
               ToolTipText     =   "Enter Customer Selection"
               Top             =   1800
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCMCDSEL"
               IdName          =   "UOCMCDSEL"
               CmpStr          =   "OmCmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1965
               TabIndex        =   15
               ToolTipText     =   "Enter From Customer"
               Top             =   1800
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
               Index           =   46
               Left            =   4680
               TabIndex        =   14
               ToolTipText     =   "Enter To Order Date"
               Top             =   1515
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
               Left            =   1965
               TabIndex        =   13
               ToolTipText     =   "Enter From Order Date"
               Top             =   1515
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
               Left            =   5295
               TabIndex        =   8
               ToolTipText     =   "Enter To Order Year"
               Top             =   945
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
               Left            =   2580
               TabIndex        =   4
               ToolTipText     =   "Enter From Order Year"
               Top             =   945
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
               Left            =   1965
               TabIndex        =   3
               ToolTipText     =   "Enter From Order Tc"
               Top             =   945
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOOMTCFR"
               IdName          =   "UOOMTCFR"
               CmpStr          =   "OdTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   3015
               TabIndex        =   5
               ToolTipText     =   "Enter From Order Character"
               Top             =   945
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
               Left            =   4680
               TabIndex        =   7
               ToolTipText     =   "Enter To Order Tc"
               Top             =   945
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataField       =   "UOOMTCTO"
               IdName          =   "UOOMTCTO"
               CmpStr          =   "OdTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   23
               Left            =   5730
               TabIndex        =   9
               ToolTipText     =   "Enter To Order Character"
               Top             =   945
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
               Index           =   53
               Left            =   3780
               TabIndex        =   6
               ToolTipText     =   "Enter From Order Number"
               Top             =   945
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
               Left            =   6510
               TabIndex        =   10
               ToolTipText     =   "Enter To Order Number"
               Top             =   945
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
               Index           =   26
               Left            =   4680
               TabIndex        =   12
               ToolTipText     =   "Enter To Order Serial"
               Top             =   1230
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRTO"
               IdName          =   "UOODSRTO"
               CmpStr          =   "OdSr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   34
               Left            =   1965
               TabIndex        =   11
               ToolTipText     =   "Enter From Order Serial"
               Top             =   1230
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOODSRFR"
               IdName          =   "UOODSRFR"
               CmpStr          =   "OdSr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1965
               TabIndex        =   0
               ToolTipText     =   "Enter From Company Code"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "OdCoCd >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   4680
               TabIndex        =   25
               ToolTipText     =   "Enter To Sales Executive"
               Top             =   2655
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERTO"
               IdName          =   "UOSALPERTO"
               CmpStr          =   "OmCmSalPer <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1965
               TabIndex        =   24
               ToolTipText     =   "Enter From Sales Executive"
               Top             =   2655
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataField       =   "UOSALPERFR"
               IdName          =   "UOSALPERFR"
               CmpStr          =   "OmCmSalPer >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   7410
               TabIndex        =   26
               ToolTipText     =   "Enter Sales Executive Selection"
               Top             =   2655
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOSALPERSEL"
               IdName          =   "UOSALPERSEL"
               CmpStr          =   "OmCmSalPer In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   4680
               TabIndex        =   1
               ToolTipText     =   "Enter To Company Code"
               Top             =   660
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDTO"
               IdName          =   "UOCOCDTO"
               CmpStr          =   "OdCoCd <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   7410
               TabIndex        =   2
               ToolTipText     =   "Enter Company Code Selection"
               Top             =   660
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOCOCDSEL"
               IdName          =   "UOCOCDSEL"
               CmpStr          =   "OdCoCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   58
               Left            =   7410
               TabIndex        =   32
               ToolTipText     =   "Enter Cell Name Selection"
               Top             =   3225
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOODCELLSEL"
               IdName          =   "UOODCELLSEL"
               CmpStr          =   "OdCell In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   60
               Left            =   1965
               TabIndex        =   30
               ToolTipText     =   "Enter From Cell Name"
               Top             =   3225
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODCELLFR"
               IdName          =   "UOODCELLFR"
               CmpStr          =   "OdCell >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4680
               TabIndex        =   28
               ToolTipText     =   "Enter To Priority Code"
               Top             =   2940
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDTO"
               IdName          =   "UOPRTCDTO"
               CmpStr          =   "OdPrtCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   7410
               TabIndex        =   29
               ToolTipText     =   "Enter Priority Code Selection"
               Top             =   2940
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UOPRTCDSEL"
               IdName          =   "UOPRTCDSEL"
               CmpStr          =   "OdPrtCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1965
               TabIndex        =   27
               ToolTipText     =   "Enter From Priority Code"
               Top             =   2940
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOPRTCDFR"
               IdName          =   "UOPRTCDFR"
               CmpStr          =   "OdPrtCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   59
               Left            =   4680
               TabIndex        =   31
               ToolTipText     =   "Enter To Cell Name"
               Top             =   3225
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOODCELLTO"
               IdName          =   "UOODCELLTO"
               CmpStr          =   "OdCell <="
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Priority"
               BeginProperty Font 
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
               Index           =   4
               Left            =   60
               TabIndex        =   60
               Top             =   2940
               Width           =   690
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
               Index           =   21
               Left            =   6450
               TabIndex        =   59
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
               Index           =   20
               Left            =   3720
               TabIndex        =   58
               Top             =   960
               Width           =   105
            End
            Begin VB.Label LblShowQty 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Shipment Status"
               BeginProperty Font 
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
               Left            =   60
               TabIndex        =   57
               Top             =   3630
               Width           =   1560
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cell"
               BeginProperty Font 
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
               Index           =   6
               Left            =   60
               TabIndex        =   55
               Top             =   3225
               Width           =   375
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive"
               BeginProperty Font 
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
               Index           =   8
               Left            =   60
               TabIndex        =   54
               Top             =   2655
               Width           =   1515
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   15
               Left            =   60
               TabIndex        =   53
               Top             =   660
               Width           =   1215
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order Sr"
               BeginProperty Font 
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
               Index           =   10
               Left            =   60
               TabIndex        =   52
               Top             =   1230
               Width           =   810
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Tc/Yy/Chr/No."
               BeginProperty Font 
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
               Index           =   22
               Left            =   60
               TabIndex        =   51
               Top             =   945
               Width           =   1680
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
               Left            =   5670
               TabIndex        =   50
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
               Index           =   25
               Left            =   2955
               TabIndex        =   49
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
               Index           =   26
               Left            =   5235
               TabIndex        =   48
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
               Index           =   27
               Left            =   2520
               TabIndex        =   47
               Top             =   960
               Width           =   105
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ord Date"
               BeginProperty Font 
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
               Index           =   19
               Left            =   60
               TabIndex        =   46
               Top             =   1515
               Width           =   840
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   17
               Left            =   60
               TabIndex        =   45
               Top             =   1800
               Width           =   900
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   14
               Left            =   60
               TabIndex        =   44
               Top             =   2370
               Width           =   960
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
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
               Height          =   240
               Index           =   13
               Left            =   60
               TabIndex        =   43
               Top             =   2085
               Width           =   1020
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
               Left            =   7410
               TabIndex        =   42
               Top             =   240
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
               Left            =   4680
               TabIndex        =   41
               Top             =   240
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
               TabIndex        =   40
               Top             =   240
               Width           =   1335
            End
         End
      End
   End
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10650
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4395
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   9600
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1005
   End
   Begin MwfCtl.MWCTL_ADC adc 
      Height          =   555
      Left            =   0
      TabIndex        =   66
      Top             =   9630
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
   Begin VB.Frame fra_keys 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4605
      TabIndex        =   63
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   64
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
         TabIndex        =   65
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
End
Attribute VB_Name = "EmrFrpShpAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The purpose of this report is to calculate the average delay
' in the shipment for the orders selected in the scope

Option Explicit
Dim Rep As New EmrRepShpAna
Dim moCn As MwfLib.MDOConnection
Private Sub SetGroupSort()
  GRP_REP.Add "Company Code", "OdCoCd", "", "", "", "", "", "", "(Select HName from Head where HCd= '" + ctSelfCmCd + "' and HCoCd=OdCoCd) "
  GRP_REP.Add "Order No", "OmCoCd+OmTc+OmYy+OmChr+str(OmNo)", "OmCoCd+'/'+OmTc+'/'+OmYy+'/'+OmChr+'/'+LTrim(Str(OmNo))", "", "", "Company Code,Customer Code,Sales Exec", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  GRP_REP.Add "Order Sr", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+str(OdNo)+'/'+Str(OdSr)", "OdCoCd+'/'+OdTc+'/'+OdYy+'/'+OdChr+'/'+LTrim(Str(OdNo))+'/'+LTrim(Str(OdSr))", "", "", "Order No,Company Code,Customer Code,Sales Exec", "", "", "'('+OmCmCd+ '/'+ OmPONo+ ')' "
  GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select (case when (Select PNum from Param where Ptyp = 'USR' and PmCd = '" + Trim(EmrFrmMDI.TXT_UCD.text) + "')  <= 5 then  CmName else CmCd end) From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
  GRP_REP.Add "Design Ctg", "DmCtg", "", "DmCtg", "hDmCtg", "", "", "", "(Select PDesc from Param where PTyp= 'DMCTG' and PMCd= DmCtg) "
  GRP_REP.Add "Design Code", "OdDmCd", "", "OdDmCd", "hOdDmCd", "Design Ctg", "", "", "(Select Substring(DmDesc,1,120) From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey='" + ctCurrPrtn + "' and ", "") + " DmTcTyp= OdDmTcTyp and DmCd= OdDmCd and DmSz= '') "
  GRP_REP.Add "Sales Exec", "OmCmSalPer", "", "", "", "", "", "", "(Select PDesc from Param Where PTyp = 'SP' and PMCd= OmCmSalPer) "
  GRP_REP.Add "Priority", "OdPrtCd", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd) "
  GRP_REP.Add "Cell", "OdCell", "", "", "", "", "", "", "(Select PDesc from Param where PTyp= 'CELL' and PMCd= OdCell) "
End Sub

Private Sub adc_setreprecsource()
  On Error GoTo RepErr
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String

  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  Dim wGrp(3) As String
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)

  wrepcnd = adc.RepCond
  
  If OptShp(0).Value = True Then
    wCnd = wrepcnd
  Else
    wCnd = wrepcnd + " and OdCls='Y' and OdOrdQty= OdExpQty "
  End If

wSqlStrg = "Select " + grpflds + ", " + _
             "OmCoCd, OmTc, OmYy, OmChr, OmNo, OdSr, OmCmCd, OmOrgExpDelDt, OdExpDelDt," + _
             "OdOrdQty, InDt, IdQty ,InCoCd, InTc, InYy, InChr, InNo, IdSr, " + _
             "case when datediff(DD,OmOrgExpDelDt,InDt)>0 and OmOrgExpDelDt<> '01/01/80' then datediff(DD,OmOrgExpDelDt,InDt) else 0 End DlywrtOrg," + _
             "case when datediff(DD,OdExpDelDt,InDt)>0 and OdExpDelDt<> '01/01/80' then datediff(DD,OdExpDelDt,InDt) else 0 End DlywrtRev," + _
             "InCmCd, DmCtg, OdDmCd ,OmCoCd+'/'+OmTc+'/'+OmYy+'/'+OmChr+'/'+ltrim(str(OmNo))+'/'+ltrim(str(OdSr)) as OrdSr " + _
             "from OrdMst " + _
             "join OrdDsg on OdOmIdNo=OmIdNo and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo and OdCoCd=OmCoCd and OdPrtKey=OmPrtKey " + _
             "join DsgMst on DmIdNo=OdDmIdNo and DmTcTyp= 'DM' and DmCd = OdDmCd " + _
             "join CustMst on CmCtg=OmCmCtg and CmCd=OmCmCd " + _
             "left outer join InvDsg on IdOdIdNo=OdIdNo and IdCoCd=OdCoCd and IdExpOdTc= OdTc and IdExpOdYy= OdYy and IdExpOdChr= OdChr and IdExpOdNo = OdNo and IdExpOdSr = OdSr " + _
             "left outer join InvHd on InIdNo= IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo and InPrtKey= IdPrtKey " + _
             "where " + wCnd + _
             " order by OmCoCd,OmTc,OmYy,OmChr,OmNo,OdSr"
          
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc)
  Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UOOMTCFR") = gs_OmTcTyp
  adc("UOOMTCTO") = gs_OmTcTyp
End Sub

Private Sub Form_GotFocus()
  SetMDISysColorsAfterLogin
End Sub
Private Sub Form_Load()
  Dim ws_BtnStr As String
  ws_BtnStr = ""
  Call SetObjColors(Me, ws_BtnStr)
End Sub
Private Sub ADC_Load()
  
  Set moCn = adc.Connection
  Set adc.FirNKeyCtl = adc("UoCoCdFr")
  Call InitProp(Me)
  Call SetGroupSort
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  
 Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  SetProp Me, IdName, When
End Sub

Private Sub OptShp_GotFocus(Index As Integer)
  DispMsg OptShp(Index).ToolTipText, etInfo
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

  Dim ws_HName As String
    ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UOCOCDFR") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
    Rep.wCoCd.SetText ws_HName
 

  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If

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
    Set moCn = Nothing
    Set Rep = Nothing
End Sub
