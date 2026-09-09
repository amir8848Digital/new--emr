VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpPckLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing List"
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
      Left            =   12540
      TabIndex        =   14
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
      TabIndex        =   13
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
      Left            =   6645
      TabIndex        =   23
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   24
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
         TabIndex        =   25
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
      Left            =   7860
      TabIndex        =   22
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   19
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
      TabIndex        =   15
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
      Height          =   9825
      Left            =   0
      TabIndex        =   18
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   17
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpPckLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpPckLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   21
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1920
               TabIndex        =   41
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   2400
               Width           =   3465
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
                  TabIndex        =   44
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
                  TabIndex        =   43
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   945
               End
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
                  TabIndex        =   42
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Box Det        "
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
               TabIndex        =   10
               Tag             =   "ShowBoxDet"
               ToolTipText     =   "Check To Show Box Details"
               Top             =   1830
               Width           =   2100
            End
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show Summary     "
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
               Tag             =   "ShowSumm"
               ToolTipText     =   "Check To Show Summary"
               Top             =   2115
               Width           =   2100
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   14
               Left            =   2340
               TabIndex        =   9
               ToolTipText     =   "Show Box Detail (Yes / No)"
               Top             =   1830
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
               Index           =   20
               Left            =   2730
               TabIndex        =   2
               ToolTipText     =   "Enter From Invoice Year"
               Top             =   975
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "IfYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   3150
               TabIndex        =   3
               ToolTipText     =   "Enter From Invoice Character"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "IfChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   26
               Left            =   3930
               TabIndex        =   4
               ToolTipText     =   "Enter From Invoice Number"
               Top             =   975
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "IfNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1950
               TabIndex        =   1
               ToolTipText     =   "Enter From Invoice Tc"
               Top             =   975
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "IfTc ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   17
               Left            =   1950
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   690
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "IfCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1950
               TabIndex        =   5
               ToolTipText     =   "Enter From Export Number"
               Top             =   1260
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   503
               MaxLength       =   20
               DataField       =   "UOINEXPNOFR"
               IdName          =   "UOINEXPNOFR"
               CmpStr          =   "InExpNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   36
               Left            =   1950
               TabIndex        =   6
               ToolTipText     =   "Enter From Packing List Year"
               Top             =   1545
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UOFDPLYYFR"
               IdName          =   "UOFDPLYYFR"
               CmpStr          =   "FdPlYy ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   37
               Left            =   2355
               TabIndex        =   7
               ToolTipText     =   "Enter From Packing List Character"
               Top             =   1545
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOFDPLCHRFR"
               IdName          =   "UOFDPLCHRFR"
               CmpStr          =   "FdPlChr ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   39
               Left            =   3135
               TabIndex        =   8
               ToolTipText     =   "Enter From Packing List Number"
               Top             =   1545
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOFDPLNOFR"
               IdName          =   "UOFDPLNOFR"
               CmpStr          =   "FdPlNo ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   2340
               TabIndex        =   11
               ToolTipText     =   "Show Summary (Yes / No)"
               Top             =   2115
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
               Left            =   5640
               TabIndex        =   39
               ToolTipText     =   "Show Open Bags (Yes / No / All)"
               Top             =   2400
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN2"
               IdName          =   "UOYN2"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Opn Bags"
               BeginProperty Font 
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
               TabIndex        =   40
               Top             =   2415
               Width           =   1665
            End
            Begin VB.Label ALBL 
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
               Index           =   3
               Left            =   90
               TabIndex        =   38
               ToolTipText     =   "Location"
               Top             =   2130
               Visible         =   0   'False
               Width           =   1755
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
               Index           =   39
               Left            =   2295
               TabIndex        =   37
               Top             =   1545
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
               Index           =   37
               Left            =   3075
               TabIndex        =   36
               Top             =   1545
               Width           =   105
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
               TabIndex        =   35
               Top             =   1260
               Width           =   1425
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Packing List"
               BeginProperty Font 
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
               Left            =   90
               TabIndex        =   34
               Top             =   1545
               Width           =   1905
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
               TabIndex        =   33
               Top             =   690
               Width           =   1335
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
               Left            =   1950
               TabIndex        =   32
               Top             =   180
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
               Left            =   4290
               TabIndex        =   31
               Top             =   180
               Width           =   1275
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
               Left            =   2670
               TabIndex        =   30
               Top             =   975
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
               Left            =   3870
               TabIndex        =   29
               Top             =   975
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
               Left            =   3090
               TabIndex        =   28
               Top             =   975
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
               TabIndex        =   27
               Top             =   975
               Width           =   1815
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
               Index           =   0
               Left            =   90
               TabIndex        =   26
               ToolTipText     =   "Location"
               Top             =   1830
               Visible         =   0   'False
               Width           =   1485
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpPckLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepPckLst
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  
  'GRP_REP.Add "Group1", "", "", "", "", "", "", "", ""
  'GRP_REP.Add "Group2", "", "", "", "", "", "", "", ""
  'GRP_REP.Add "Group3", "", "", "", "", "", "", "", ""
    
  GRP_REP.Add "Design Ctg", "DmCtg", "", "", "", "Sales Ctg", "", "", "(Select PDesc From Param Where PTyp= 'DMCTG' and PmCd= DmCtg) "
  'Geeta***Crm***212
  GRP_REP.Add "Sales Ctg", "DmSalCtg", "", "", "", "Design Ctg", "", "", "(Select PDesc From Param Where PTyp= 'SALCTG' and Pmcd=DmCtg and PSCd= DmSalCtg) "
  GRP_REP.Add "Design Code", "FdDmCd", "", "FdDmCd", "hFdDmCd", "Design Ctg,Sales Ctg", "", "", "(Select DmDesc From DsgMst Where " + IIF(gs_Partition = ctCurrPrtn, " DmPrtKey=FdPrtKey and ", "") + " DmTcTyp= FdDmTcTyp and DmCd= FdDmCd and DmSz= '') "
  GRP_REP.Add "(DsgCd)Colour", "FdDmCd+OdDmCol", "'('+FdDmCd+')'+OdDmCol", "FdDmCd", "hFdDmCd", "Design Ctg,Sales Ctg,Design Code,(DsgCd)Suffix,(DsgCd)Size", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "(DsgCd)Suffix", "FdDmCd+FdSfx", "'('+FdDmCd+')'+FdSfx", "FdDmCd,FdSfx", "hFdDmCd,hFdSfx", "", "", "", ""
  GRP_REP.Add "(DsgCd)Size", "FdDmCd+FdDmSz", "'('+FdDmCd+')'+FdDmSz", "FdDmCd,FdDmSz", "hFdDmCd,hFdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PMCd = FdDmSz) "
  GRP_REP.Add "Design Size", "FdDmSz", "", "FdDmSz", "hFdDmSz", "", "", "", "(Select PDesc From Param Where PTyp= 'DMSZ' and PMCd = FdDmSz) "
  GRP_REP.Add "Design Colour", "OdDmCol", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMCOL' and PMCd = OdDmCol) "
  GRP_REP.Add "PCS/PRS", "DmUom", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'DMUOM' and PmCd= DmUom) "
  GRP_REP.Add "Ord Karat", "OdKt", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'KT' and PMCd = OdKt) "
  GRP_REP.Add "Ord VaCtg", "OdVaCtg", "", "", "", "", "", "", "(Select PDesc From Param Where PTyp= 'VACTG' and PMCd = OdVaCtg) "
  
  GRP_REP.Add "ExpOrderNo", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)", "FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))", "", "", "Exp Customer", "", "", "'('+FdExpCmCd+ ')' "
  GRP_REP.Add "Exp Customer", "FdExpCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= FdExpCmCd) "
  GRP_REP.Add "ExpOrderSr", "FdExpOdTc+FdExpOdYy+FdExpOdChr+str(FdExpOdNo)+str(FdExpOdSr)", "FdExpOdTc+'/'+FdExpOdYy+'/'+FdExpOdChr+'/'+LTrim(Str(FdExpOdNo))+'/'+LTrim(Str(FdExpOdSr))", "wOrdSr", "hOrdSr", "Exp Customer,ExpOrderNo,Prd Customer,PrdOrderNo", "", "", "'('+FdExpCmCd+ ')' "
  GRP_REP.Add "PrdOrderNo", "FdPrdOdTc+FdPrdOdYy+FdPrdOdChr+str(FdPrdOdNo)", "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo))", "", "", "Prd Customer", "", "", "'('+FdPrdCmCd+ ')' "
  GRP_REP.Add "Prd Customer", "FdPrdCmCd", "", "", "", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= FdPrdCmCd) "
  GRP_REP.Add "PrdOrderSr", "FdPrdOdTc+FdPrdOdYy+FdPrdOdChr+str(FdPrdOdNo)+str(FdPrdOdSr)", "FdPrdOdTc+'/'+FdPrdOdYy+'/'+FdPrdOdChr+'/'+LTrim(Str(FdPrdOdNo))+'/'+LTrim(Str(FdPrdOdSr))", "", "", "Prd Customer,PrdOrderNo,Exp Customer,ExpOrderNo,", "", "", "'('+FdPrdCmCd+ ')' "
  'GRP_REP.Add "InvoiceNo", "IfTc+IfYy+IfChr+str(IfNo)", "IfTc+'/'+IfYy+'/'+IfChr+'/'+LTrim(Str(IfNo))", "", "", "", "", "", "(Select InExpNo from InvHd where InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo) "
    
  GRP_REP.Add "(DsgCd)InvSr", "FdDmCd+str(IfSr)", "FdDmCd+' ('+LTrim(str(IfSr))+')'", "FdDmCd", "hFdDmCd", "Design Ctg,Sales Ctg,Design Code,(DsgCd)Colour,(DsgCd)Suffix,(DsgCd)Size", "", "", ""
  GRP_REP.Add "Cust PONo(Exp)", "OmPoNo", "", "", "", "", "", "", ""
  
    
'  Group options like DmCd, (DmCd)+Sfx, (DmCd)+Sz, (DmCd)+OdDmCol, OdKt, ExpOrdNo, ExpCmCd, PrdOrdNo, PrdCmCd, and as u feel reqd.
'If u want that the order of Packing list be same as other Invoice documents, then this is a little tricky - but worth a Try.

'DmCtg, DmSalCtg, FdDmCd, FdSfx, FdDmSz, OdDmCol, Pcs/Prs, OdKt, OdVaCtg, ExpOrdNo, ExpOrdSr,
'FdExpCmCd, PrdOrdNo, PrdOrdSr, FdPrdCmCd, InvNo,
End Sub
Private Sub adc_setreprecsource()
  On Error GoTo errhld    ' ****** Manali 3.5.0 - 15/06/09
  
  '*** Report Sql ***
  'Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  'space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  'space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  'FdTc, FdYy, FdChr, FdNo, FdSr, FdDmCd, FdDmSz, FdSfx, FdBYy, FdBChr,
  'FdBNo, FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, FdExpOdSr, IfQty,
  'IfGrsWt, FrRmWt as qNetWt, FrRmCd, FrRmSz, FrRmStkRt, PValue3 as qPtrYN,
  'FrRmQty, FrRmWt
  'From InvFgd, Fgd, OrdDsg, DsgMst, FgRm, RmMst, Param where 1= 2
  '*** Report Sql ***
    
  'GRP_REP.Value(gltGroup1) = "Grp2"
  'GRP_REP.Value(gltGroup2) = "(None)"
 
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wInvDt As String
  Dim wTcTypCnd As String, wRmSummSql As String
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  
  '4.1.0.0
  If adc("UOYN2") = "Y" Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " IfSr=0 "
  ElseIf adc("UOYN2") = "N" Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " IfSr<>0 "
  End If
  
  '*** To check if Detail is selected ***
  Dim wGrp(6) As String, i As Integer, wDetPos As Integer
  wGrp(1) = GRP_REP.Value(gltGroup1)
  wGrp(2) = GRP_REP.Value(gltGroup2)
  wGrp(3) = GRP_REP.Value(gltGroup3)
  wDetPos = 0
  For i = 1 To 6
    If wGrp(i) = "(Detail)" Then wDetPos = i: Exit For
  Next i
  
  '*** If Detail not selected then suppress Box Header, Detail and Footer ***
  If wDetPos = 0 Then
    Rep.BoxHd.Suppress = True: Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True: Rep.DetSec2.Suppress = True
  End If
  
  If UCase(adc.MenuCd) = UCase("RepPckLst") Then
    Rep.TxtHead.SetText "Finished Goods Packing List For INV"
    Rep.TxnVchNo.SetText "Invoice Vch No: "
    wTcTypCnd = " where (Select PValue from Param where Ptyp= 'TC' and PMCd= IfTc)= 'IN' "
  ElseIf UCase(adc.MenuCd) = UCase("RepCTBPckLst") Then
    Rep.TxtHead.SetText "Finished Goods Packing List For CTB"
    Rep.TxnVchNo.SetText "CTB Vch No: "
    wTcTypCnd = " where (Select PValue from Param where Ptyp= 'TC' and PMCd= IfTc)= 'CTB' "
  End If
  
  '****** Sachin 3.02 - Id Fields in Join
  ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
  wSqlStrg = "Select " + grpflds + ", " + _
             " FdTc, FdYy, FdChr, FdNo, FdSr, FdDmCd, FdDmSz, FdSfx, FdBYy, FdBChr, FdBNo, " + _
             "FdExpOdTc, FdExpOdYy, FdExpOdChr, FdExpOdNo, FdExpOdSr, IfQty, IfGrsWt, " + _
             "(case when RmCtg in ('G', 'P', 'S', 'L','X') then FrRmWt else 0 end) as qNetWt, FrRmCd, " + _
             "FrRmSz, FrRmSz2, FrRmSz3, FrRmStkRt, " + _
             "(Select PValue3 from Param where PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg) as qPtrYN, " + _
             "FrRmQty, FrRmWt " + _
             "From InvFgd join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy " + _
             "and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=IfPrtKey ", "") + _
             "join OrdDsg on OdCoCd= FdCoCd and OdTc= FdExpOdTc and OdYy= FdExpOdYy " + _
             "and OdChr= FdExpOdChr and OdNo= FdExpOdNo and OdSr= FdExpOdSr " + IIF(gs_Partition = ctCurrPrtn, " and OdPrtKey=FdPrtKey ", "") + _
             "join DsgMst on DmIdNo=FdDmIdNo and DmTcTyp= FdDmTcTyp and DmCd= FdDmCd and DmSz= '' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=FdPrtKey ", "") + _
             "join OrdMst on OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy " + _
             "and OmChr= OdChr and OmNo= OdNo " + IIF(gs_Partition = ctCurrPrtn, " and OmPrtKey=OdPrtKey ", "") + _
             "join InvHd on InIdNo=IfInIdNo and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and " + _
             "InChr= IfChr and InNo= IfNo " + IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IfPrtKey ", "") + _
             "join FgRm On FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr " + _
             "and FrNo= FdNo and FrSr= FdSr " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=OdPrtKey ", "") + _
             "join RmMst on RmCd= FrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + wTcTypCnd + _
             wCnd + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey='" + ctCurrPrtn + "' ", "")
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  
  '*** Jay 3.1.0_1 [RmCtg SubReport]
  If adc("UoYN1") = "Y" Then
    Rep.RF2.Suppress = False
    
    '****** Sachin 3.02 - If Fields In Join
    ' ***** Manali 3.03 - 24/06/08 - Pd Mod - 'L' added
    ' ***** Manali - 15/06/09 - Error rectified - Join with InvHd added
    wRmSummSql = "Select RmCtg, Max(PDesc) as PDesc, Sum(FrRmQty) as FrRmQty, Sum(FrRmWt) as FrRmWt, " + _
        "sum(qPureWt) As qPureWt " + _
        "From ( Select RmCtg, PDesc, FrRmQty, FrRmWt, " + _
        "        FrRmWt* (Case when RmCtg in ('G', 'P', 'S', 'L') then RmPurityWt when RmCtg= 'X' then RmBasePurityWt" + _
        "          else 0 end) /(case when Rm.RmCtg in ('G', 'P', 'S', 'L')    then IsNull((Select R1.RmPurityWt from Param " + _
        "          join RmMst R1 on " + IIF(gs_Partition = ctCurrPrtn, " R1.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " R1.RmCd= PValue where PTyp= 'RMCTG' and PMCd= Rm.RmCtg), 1) when Rm.RmCtg= 'X' then " + _
        "          IsNull((Select R2.RmPurityWt from Param join RmMst R2 on " + IIF(gs_Partition = ctCurrPrtn, " R2.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " R2.RmCd= PValue where PTyp= 'RMCTG' " + _
        "          and PMCd= IsNull((Select R3.RmCtg From RmMst R3 where " + IIF(gs_Partition = ctCurrPrtn, " R3.RmPrtKey='" + ctCurrPrtn + "' and ", "") + " R3.RmCd= Rm.RmBaseCd), '')), 1) else 1 end) as qPureWt " + _
        "         From InvFgd Join InvHd On InIdNo=IfInIdNo and InCoCd=IfCoCd and InTc=IfTc and InYy=IfYy and InChr=IfChr and InNo=IfNo " + _
                      IIF(gs_Partition = ctCurrPrtn, " and InPrtKey=IfPrtKey ", "") + _
        "      join Fgd on FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo " + _
        "        and FdSr= IfFdSr " + IIF(gs_Partition = ctCurrPrtn, " and FdPrtKey=IfPrtKey ", "") + _
        "      join FgRm On FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr " + IIF(gs_Partition = ctCurrPrtn, " and FrPrtKey=FdPrtKey ", "") + _
        "      join RmMst Rm on RmCd= FrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and RmPrtKey=FrPrtKey ", "") + _
        "      Join Param On PTyp= 'RMCTG' and PMCD= RMCTG " + wTcTypCnd + wCnd + IIF(gs_Partition = ctCurrPrtn, " and IfPrtKey='" + ctCurrPrtn + "' ", "") + _
        "    ) qTbl " + _
        "Group By RmCtg "
       
    Call moCn.RepRes(Rep.Subreport1.OpenSubreport, wRmSummSql + ctMaxDopOpt)   '****** Sachin 3.01 added ctMaxDopOpt
  Else
    Rep.RF2.Suppress = True
  End If
  '*** Jay 3.1.0_1 [RmCtg SubReport]

  If adc("UoYN") = "N" Then
    Rep.BoxHd.Suppress = True
    Rep.BoxDet.Suppress = True
    Rep.BoxFt.Suppress = True
    Rep.DetSec2.Suppress = True
  End If
  
  wInvDt = moCn.GetFldVal("Select InDt From InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd = '" + adc("UoCoCdFr") + "' and " + _
                        IIF(adc("UoInExpNoFr") = "", "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
                        "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " ", " InExpNo= '" + CStr(adc("UoInExpNoFr")) + "' "))
  SetRepText Rep, "TxtInvNo", IIF(adc("UoInExpNoFr") = "", Trim(adc("UoInTcFr")) + "/" + Trim(adc("UoInYyFr")) + "/" + Trim(adc("UoInChrFr")) + "/" + Trim(CStr(adc("UoInNoFr"))), _
                              moCn.GetFldVal("Select InTc+'/'+InYy+'/'+InChr+'/'+LTrim(Str(InNo)) From InvHd  where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd = '" + adc("UoCoCdFr") + "' and InExpNo= '" + CStr(adc("UoInExpNoFr")) + "' "))
  
'  wInvDt = moCn.GetFldVal("Select InDt From InvHd where " + IIF(gs_Partition = ctCurrPrtn, " InPrtKey='" + ctCurrPrtn + "' and ", "") + " InCoCd = '" + adc("UoCoCdFr") + "' and " + _
'                        "InTc= '" + adc("UoInTcFr") + "' and InYy= '" + adc("UoInYyFr") + "' and " + _
'                        "InChr= '" + adc("UoInChrFr") + "' and InNo= " + CStr(adc("UoInNoFr")) + " ")
  'SetRepText Rep, "TxtInvNo", Trim(adc("UoInTcFr")) + "/" + Trim(adc("UoInYyFr")) + "/" + Trim(adc("UoInChrFr")) + "/" + Trim(CStr(adc("UoInNoFr")))
  '6.1
  'SetRepText Rep, "TxtInvDt", Format(IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wInvDt, "mm") + "/" + Format(wInvDt, "dd"), Format(wInvDt, "dd") + "/" + Format(wInvDt, "mm")) + "/" + Format(wInvDt, "yyyy"), DtShortStrYyyy())
  SetRepText Rep, "TxtInvDt", IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", Format(wInvDt, "mm") + "/" + Format(wInvDt, "dd"), Format(wInvDt, "dd") + "/" + Format(wInvDt, "mm")) + "/" + Format(wInvDt, "yyyy")
  
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.FrRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
Exit Sub

errhld:
     Call ChkSelScope(Err.Number, Err.Description)
  
End Sub

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_Click(Index As Integer)
With ChkBoxArr(Index)
  Select Case Index
  Case Is = 0     'Tag = "ShowBoxDet"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  Case Is = 1     'Tag = "ShowSumm"
    If .Value = Checked Then adc("UoYN1") = "Y" Else adc("UoYN1") = "N"
  End Select
End With
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added

' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN2") = "Y"
Case Is = 1
  adc("UoYN2") = "N"
Case Is = 2
  adc("UoYN2") = ""
End Select
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
    Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    If UCase(adc.MenuCd) = UCase("RepPckLst") Then
      gs_InTcTyp = "IN"
    ElseIf UCase(adc.MenuCd) = UCase("RepCTBPckLst") Then
      gs_InTcTyp = "CTB"
    End If
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
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
 
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  If adc("UoYN") = "" Then adc("UoYN") = "N"
  If adc("UoYN1") = "" Then adc("UoYN1") = "N"
  If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
  If adc("UoYN1") = "Y" Then ChkBoxArr(1).Value = Checked Else ChkBoxArr(1).Value = Unchecked
  ' ****** Manali 3.5.0 - 19/11/08 - Check Boxes added
  
  '4.1.0.0
  If adc("UoYN2") = "Y" Then
    OptBag(0).Value = True
  ElseIf adc("UoYN2") = "N" Then
    OptBag(1).Value = True
  ElseIf adc("UoYN2") = "" Then
    OptBag(2).Value = True
  End If

End Sub
Private Sub DispCoNm()
'Rep.wCoCd.SetText gs_CoNm
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
