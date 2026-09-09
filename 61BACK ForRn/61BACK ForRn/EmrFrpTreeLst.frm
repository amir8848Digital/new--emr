VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpTreeLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Listing"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_HLP MWCTL_HLP1 
      Height          =   495
      Left            =   10920
      TabIndex        =   20
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
      TabIndex        =   21
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
      Left            =   4875
      TabIndex        =   30
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   31
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
         TabIndex        =   32
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
      Left            =   6090
      TabIndex        =   29
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   26
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
      TabIndex        =   22
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
      TabIndex        =   25
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   24
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpTreeLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpTreeLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   28
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   23
               Top             =   360
               Width           =   9585
               _ExtentX        =   16907
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   360
            Width           =   15105
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
               Left            =   30
               TabIndex        =   15
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   2520
               Width           =   1680
            End
            Begin VB.Frame FraOptTree 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1515
               TabIndex        =   46
               ToolTipText     =   "Show Casted Or Non-Casted Or All Tree Details"
               Top             =   2805
               Width           =   4095
               Begin VB.OptionButton OptTree 
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
                  Left            =   3060
                  TabIndex        =   19
                  ToolTipText     =   "Show Casted Or Non-Casted Or All Tree Details"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptTree 
                  Caption         =   "Non-Casted"
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
                  TabIndex        =   18
                  ToolTipText     =   "Show Casted Or Non-Casted Or All Tree Details"
                  Top             =   0
                  Width           =   1605
               End
               Begin VB.OptionButton OptTree 
                  Caption         =   "Casted"
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
                  TabIndex        =   17
                  ToolTipText     =   "Show Casted Or Non-Casted Or All Tree Details"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4005
               TabIndex        =   8
               ToolTipText     =   "Enter To Tree  Date"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "TrDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   48
               Left            =   1515
               TabIndex        =   7
               ToolTipText     =   "Enter FromTree Date"
               Top             =   1665
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "TrDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   19
               Left            =   1515
               TabIndex        =   1
               ToolTipText     =   "Enter From Tree YyMmDd"
               Top             =   1095
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   6
               DataType        =   4
               DataField       =   "UOTRYMDFR"
               IdName          =   "UOTRYMDFR"
               CmpStr          =   "trYmd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   49
               Left            =   2430
               TabIndex        =   2
               ToolTipText     =   "Enter From Tree Character"
               Top             =   1095
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
               CmpStr          =   "TrChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   50
               Left            =   4005
               TabIndex        =   3
               ToolTipText     =   "Enter To Tree YyMmDd"
               Top             =   1095
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   6
               DataType        =   4
               DataField       =   "UOTRYMDTO"
               IdName          =   "UOTRYMDTO"
               CmpStr          =   "TrYmd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   51
               Left            =   4905
               TabIndex        =   4
               ToolTipText     =   "Enter To Tree Character"
               Top             =   1095
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOPMCDTO"
               IdName          =   "UOPMCDTO"
               CmpStr          =   "TrChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   25
               Left            =   1515
               TabIndex        =   5
               ToolTipText     =   "Enter From Tree Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTRNOFR"
               IdName          =   "UOTRNOFR"
               CmpStr          =   "TrNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   55
               Left            =   4005
               TabIndex        =   6
               ToolTipText     =   "Enter To Tree Number"
               Top             =   1380
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               Alignment       =   2
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOTRNOTO"
               IdName          =   "UOTRNOTO"
               CmpStr          =   "TrNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   4005
               TabIndex        =   12
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   2235
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "TrRmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6510
               TabIndex        =   13
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   2235
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "TrRmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1515
               TabIndex        =   11
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   2235
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   503
               MaxLength       =   12
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               CmpStr          =   "TrRmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   1515
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   810
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "TrCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   2
               Left            =   1860
               TabIndex        =   14
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   2520
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
               Index           =   14
               Left            =   6000
               TabIndex        =   16
               ToolTipText     =   "Show  Casted/Non Casted/All  Tree Details (Yes / No/'')"
               Top             =   2805
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
               Index           =   1
               Left            =   4005
               TabIndex        =   10
               ToolTipText     =   "Enter To Casting  Date"
               Top             =   1950
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTTO"
               IdName          =   "UODELDTTO"
               CmpStr          =   "TrCstDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1515
               TabIndex        =   9
               ToolTipText     =   "Enter From Casting Date"
               Top             =   1950
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODELDTFR"
               IdName          =   "UODELDTFR"
               CmpStr          =   "TrCstDt >="
            End
            Begin VB.Label LblCstDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Cast Date"
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
               TabIndex        =   45
               Top             =   1950
               Width           =   1425
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
               Index           =   4
               Left            =   60
               TabIndex        =   44
               ToolTipText     =   "Location"
               Top             =   2520
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.Label LblTrDet 
               BackStyle       =   0  'Transparent
               Caption         =   "Tree Details"
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
               TabIndex        =   43
               Top             =   2805
               Width           =   1485
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
               Index           =   0
               Left            =   60
               TabIndex        =   42
               Top             =   810
               Width           =   1335
            End
            Begin VB.Label LblRmCd 
               BackStyle       =   0  'Transparent
               Caption         =   "Tree KT"
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
               TabIndex        =   41
               Top             =   2235
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tree Ymd/Chr"
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
               Left            =   60
               TabIndex        =   40
               Top             =   1095
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tree No"
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
               Left            =   60
               TabIndex        =   39
               Top             =   1380
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
               Index           =   28
               Left            =   4845
               TabIndex        =   38
               Top             =   1095
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
               Left            =   2355
               TabIndex        =   37
               Top             =   1095
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Tree Date"
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
               TabIndex        =   36
               Top             =   1665
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
               Left            =   6510
               TabIndex        =   35
               Top             =   180
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
               Left            =   4005
               TabIndex        =   34
               Top             =   180
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
               Left            =   1515
               TabIndex        =   33
               Top             =   180
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpTreeLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep1 As New EmrRepCstTreeLst  'Tree performance
Dim Rep2 As New EmrRepTreeLst     'Tree List
Dim Rep As CRAXDRT.Report
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Dim mTreeTyp As String

'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******
Dim ws_MultiCoMod As Boolean, ws_FormCaption As String
'****** Sachin 2.14.0 - [25-11-2006] - [Multi-Company] ******

Private Sub SetGroupSort()
'6.1 Tree date, casting date
  GRP_REP.Add "Tree Date", "convert(char(8),TrDt,112)", "convert(varchar, TrDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "TrDt", "hTrDt", "", "", "", ""
  If mTreeTyp = "CL" Then _
    GRP_REP.Add "Casting Date", "convert(char(8),TrCstDt,112)", "convert(varchar, TrCstDt, " + CStr(DtFmtSlahYyyyVal()) + ")", "TrCstDt", "hTrDt", "", "", "", ""
  GRP_REP.Add "Tree Kt", "TrRmCd", "TrRmCd", "TrRmCd", "hKt", "", "", "", "(select RmDesc from RmMst where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " RmCd=TrRmCd)"
  GRP_REP.Add "Tree Color", "TrCol", "TrCol", "TrCol", "hTrCol", "", "", "", "(select PDesc from Param where PTyp='DMCOl' and PmCd=TrCol) "
End Sub
Private Sub adc_setreprecsource()
  
  '*** Tree Investment Report Sql ***
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(4) as wTrTyp, TrCoCd,TrYmd,TrChr,TrNo,TrDt,TrRmCd,TrCol,
'  TrWaxSpGr, TrGrsWaxWt,TrNetMetWt,TrAddMetWt, TrNetMetWt as PureWt,
'  TrdSr, TrdBYy, TrdBChr, TrdBNo, TrdBagPcs, TrdDcWt, OdTc,
'  OdYy , OdChr, OdNo, OdSr, OdDmCd, OdDmCol, OdKt, OmCmCd
'  From TreeMst, TreeDet, Bag, OrdDsg, OrdMst
'  Where 1 = 2

  '***  Tree Investment Report Sql ***
  
  '***  Tree Performance Report Sql ***
  ' ***** Manali 3.6.0 - 02/09/09 - TrGrsWaxWt, TrNetMetWt, TrAddMetWt  added below
'  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'  space(4) as wTrTyp, TrCoCd,TrYmd,TrChr,TrNo,TrDt,TrRmCd,TrCol,TrGrsWaxWt, TrNetMetWt, TrAddMetWt,
'  TrCstDt,TrFrMetWt,TrUsdMetWt,TrRecdMetWt,TrCastLoss,TrCutLoss,TrScrap,TrGrindLoss,
'  TrStripLoss, TrGrsPcWt,
'  TrdSr, TrdBYy, TrdBChr, TrdBNo, TrdBagPcs, TrdDcWt, OdTc,
'  OdYy , OdChr, OdNo, OdSr, OdDmCd, OdDmCol, OdKt, OmCmCd
'  From TreeMst, TreeDet, Bag, OrdDsg, OrdMst
'  Where 1 = 2

   '***  Tree Performance Report Sql ***

  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wi_Opt As Integer
  Dim wSqlStrg1 As String, wSqlStrg2 As String, wSqlStrg3 As String
  
  Set Rep1 = Nothing
  Set Rep2 = Nothing
  
  Select Case UCase(mTreeTyp)
  Case "CL"
    Set adc.RepSource = Rep1
    Set Rep = Rep1
  Case "TL"
    Set adc.RepSource = Rep2
    Set Rep = Rep2
  End Select
   Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  '*** To check if Grp1, Grp2 is selected and also to check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer, wNonePos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0: wNonePos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then
      wDetPos = i: Exit For
    End If
    If wGrp(i) = "(None)" Then
      wNonePos = i: Exit For
    End If
  Next i
  ' If wDetPos = 0 Then DispMsg "Detail Option Has To Be Selected For This Report", etError
 If wDetPos = 0 Then
    Rep.TrDET.Suppress = True
    Rep.DetHd.Suppress = True
    Rep.DetF.Suppress = True
    
    Rep1.G4F.Suppress = True    '*** Jenny
  End If

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " where " + wrepcnd, "")
  
  If mTreeTyp = "CL" Then
      wCnd = IIF(wCnd <> "", wCnd + " and ", wCnd) + "TrCstDt<> '01/01/80'"
  Else
    If adc("UoYN1") = "Y" Then
      wCnd = IIF(wCnd <> "", wCnd + " and ", wCnd) + "TrCstDt<> '01/01/80'"
    ElseIf adc("UoYN1") = "N" Then
      wCnd = IIF(wCnd <> "", wCnd + " and ", wCnd) + "TrCstDt= '01/01/80'"
    End If
  End If
     
   '  wSqlStrg1 = "Select " + grpflds + ", " + _
                "'" + mTreeTyp + "' as wTrTyp, TrCoCd, TrYmd, TrChr, TrNo, TrDt, " + _
                "TrRmCd, TrCol, TrWaxSpGr, TrGrsWaxWt, TrNetMetWt, " + _
                "TrAddMetWt, (TrNetMetWt+TrAddMetWt)*(select RmSpGrav from RmMst where RmCd=TrRmCd)/ " + _
                "case isnull((select RmSpGrav from RmMst where RmCd=(select Pvalue from Param join RmMst  " + _
                "on  PTyp='RMCTG' and PMCd=RmCtg where Rmcd=TrRmCd)),1) when 0 then 1 else " + _
                "isnull((select RmSpGrav from RmMst where RmCd=(select Pvalue from Param join RmMst  " + _
                "on  PTyp='RMCTG' and PMCd=RmCtg where Rmcd=TrRmCd)),1) end as PureWt, "
                
      '****** (Jen 3.2.0) replaced 'RmSpGrav' by 'RmPurityWt'
      '********Bhavna added TrDustWt 5.1.0
      wSqlStrg1 = "Select " + grpflds + ", " + _
                "'" + mTreeTyp + "' as wTrTyp, TrCoCd, TrYmd, TrChr, TrNo, TrDt, TrDustWt, " + _
                "TrRmCd, TrCol, TrWaxSpGr, TrGrsWaxWt, TrNetMetWt, " + _
                "TrAddMetWt, (TrNetMetWt+TrAddMetWt)*(select RmPurityWt from RmMst where RmCd= TrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey='" + ctCurrPrtn + "' ", "") + ")/ " + _
                "case isnull((select RmPurityWt from RmMst where RmCd=(select PValue from Param , RmMst  " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " PTyp='RMCTG' and PMCd=RmCtg and Rmcd= TrRmCd)),1) when 0 then 1 else " + _
                "isnull((select RmPurityWt from RmMst where RmCd=(select Pvalue from Param , RmMst  " + _
                "where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + " PTyp='RMCTG' and PMCd= RmCtg And Rmcd=TrRmCd)),1) end as PureWt, "
     
     ' ***** Manali 3.6.0 - 02/09/09 - TrGrsWaxWt, TrNetMetWt, TrAddMetWt  added below
     wSqlStrg2 = "Select " + grpflds + ", " + _
                "'" + mTreeTyp + "' as wTrTyp, TrCoCd, TrYmd, TrChr, TrNo, TrDt, TrGrsWaxWt, TrNetMetWt, TrAddMetWt, " + _
                "TrRmCd, TrCol, TrCstDt, TrFrMetWt, TrDustWt, " + _
                "TrUsdMetWt, TrRecdMetWt, TrCastLoss, TrCutLoss, TrScrap, " + _
                "TrGrindLoss, TrStripLoss, TrGrsPcWt, "
                     
     '****** Sachin 3.02 - Id fields in Joins
     wSqlStrg3 = "TrdSr, TrdBYy, TrdBChr, TrdBNo, TrdBagPcs, " + _
                "TrdDcWt, TrdMetXWt, TrdDesc, OdTc, OdYy, OdChr, OdNo, OdSr, OdDmCd, OdDmCol, OdKt, OmCmCd " + _
                "From TreeMst  " + _
                "Left Outer join ( TreeDet " + _
                "   join Bag on " + IIF(gs_Partition = ctCurrPrtn, " BPrtKey='" + ctCurrPrtn + "' and ", "") + " BCoCd=TrdCoCd and BYy=TrdByy and BChr=TrdBChr and BNo=TrdBNo " + _
                "   join OrdDsg on " + IIF(gs_Partition = ctCurrPrtn, " OdPrtKey=BPrtKey and ", "") + " OdIdNo=BOdIdNo and  OdCoCd=BCocd and OdTc=BOdTc and OdYy=BOdyy and OdChr=BOdChr  " + _
                "     and OdNo=BOdNo and OdSr=BOdSr " + _
                "   join OrdMst on " + IIF(gs_Partition = ctCurrPrtn, " OmPrtKey=OdPrtKey and ", "") + " OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr  " + _
                "     and OmNo=OdNo ) " + _
                "on TrCoCd=TrdCoCd and TrYmd=TrdYmd and TrChr=TrdChr and TrNo=TrdNo " + wCnd
   
   If mTreeTyp = "TL" Then
    wSqlStrg = wSqlStrg1 + wSqlStrg3 + " Order by TrCoCd,TrYmd,TrChr,TrNo,TrdSr"
   Else
     wSqlStrg = wSqlStrg2 + wSqlStrg3 + " Order by TrCoCd,TrYmd,TrChr,TrNo,TrdSr"
   End If
   adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
   If adc("UOYN") = "N" Then
    Rep.DetHd.Suppress = True
    Rep.TrDET.Suppress = True
    Rep.DetF.Suppress = True
  End If
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
End Sub

' ****** Manali 3.5.0 - 21/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
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
  
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  ws_MultiCoMod = IIF(moCn.GetFldVal("Select hImpComLoc From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + ctSelfCoCd + "'") = "Y", True, False)
  gb_CoCdFor = True
  Set adc.FirNKeyCtl = IIF(ws_MultiCoMod = True, adc("UoCoCdFr"), adc("UoTrYmdFr"))
      adc("UOCOCDFR").Enabled = ws_MultiCoMod
  '****** Sachin 2.14.0 - [23-11-2006] - [Multi-Company] ******
  
  Call InitProp(Me)
   ' mTreeTyp = "CL"
   If UCase(adc.MenuCd) = UCase("RepTreeInvst") Then mTreeTyp = "TL": Me.Caption = "Tree Investment Report "
   If UCase(adc.MenuCd) = UCase("RepTreePerform") Then mTreeTyp = "CL": Me.Caption = "Tree Performance Report"
  
  ws_FormCaption = Me.Caption  '****** Sachin 2.14.0
  
  If mTreeTyp = "CL" Then
    LblTrDet.Visible = False
    adc("UoYN1").Visible = False
    FraOptTree.Enabled = False: FraOptTree.Visible = False
    LblCstDt.Visible = True
    adc("UoDelDtFr").Visible = True
    adc("UoDelDtTo").Visible = True
  Else
    LblTrDet.Visible = True
    'adc("UoYN1").Visible = True
    FraOptTree.Enabled = True: FraOptTree.Visible = True
    LblCstDt.Visible = False
    adc("UoDelDtFr").Visible = False
    adc("UoDelDtTo").Visible = False
  End If
  Call SetGroupSort
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  adc("UoCoCdFr") = gs_CoCd
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  Me.Caption = GetFrmCaption(adc("UOCOCDFR"), ws_FormCaption)
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  If mTreeTyp = "TL" Then _
    adc("UoDelDtFr") = "01/01/80": adc("UoDelDtTo") = "01/01/80"                'MWLib.EmptyDate
    
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  If adc("UoYN1") = "Y" Then
    OptTree(0).Value = True
  ElseIf adc("UoYN1") = "N" Then
    OptTree(1).Value = True
  ElseIf adc("UoYN1") = "" Then
    OptTree(2).Value = True
  End If
  ' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
  
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
'Select Case IdName
'  Case Is <> UCase("UoRmCdFr"), UCase("UoRmCdTo"), UCase("UoRmCdSEL")

  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
  gb_CoCdFor = True
  Select Case UCase(IdName)
    Case Is = UCase("UoCoCdFr")
     If Not moCn.RecSeek("Select 'x' From Head Where HCd='" + ctSelfCmCd + "' And HCoCd='" + pv_NewValue + "'") Then _
          Cancel = True: ErrMsg = "Invalid Company Code": Exit Sub
     Me.Caption = GetFrmCaption(pv_NewValue, ws_FormCaption)
  End Select
  '****** Sachin 2.14.0 - [24-11-2006] - [Multi-Company] ******
    
    Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
    ' End Select
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_PTyp = "TRCHR"
  gb_CoCdFor = True
  
  If UCase(IdName) = UCase("UoRmCdFr") Or UCase(IdName) = UCase("UoRmCdTo") Or UCase(IdName) = UCase("UoRmCdSel") Then
      If UCase(IdName) = UCase("UoRmCdSEL") Then HlpList.MultiSelect = True
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    Call HlpList.RmCd("N", "'G','P','S', 'L'")
  Else
    SetProp Me, IdName, When
  End If
End Sub

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptTree_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
Case Is = 2
  adc("UoYN1") = ""
End Select
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added
Private Sub OptTree_GotFocus(Index As Integer)
  DispMsg FraOptTree.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 06/11/08 - Radio Buttons Added

Private Sub tab_REP_Click(PreviousTab As Integer)
  fra_tabrep(TAB_REP.Tab).Enabled = True
  fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
  adc("UOUSRCD") = adc.UsrCd
  adc("UOMNUCD") = adc.MenuCd
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm
'  Rep2.wCoCd.SetText gs_CoNm
'  Rep3.wCoCd.SetText gs_CoNm
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd ='" + adc("UoCoCdFr") + "' and " + _
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
  Rep1.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  Rep2.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm
  
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
    Set Rep1 = Nothing
    Set Rep2 = Nothing
    '*** (09/08/05)
End Sub
