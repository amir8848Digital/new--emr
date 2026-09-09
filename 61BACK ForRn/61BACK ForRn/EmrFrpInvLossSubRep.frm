VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpInvLossSubRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Metal Loss Sub Report"
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
      Left            =   11160
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   18
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
      Left            =   5325
      TabIndex        =   26
      Top             =   9600
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   27
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
         TabIndex        =   28
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
      Left            =   6540
      TabIndex        =   25
      Top             =   9615
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   23
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
      TabIndex        =   17
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
      TabIndex        =   22
      Top             =   60
      Width           =   15255
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   21
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpInvLossSubRep.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpInvLossSubRep.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraVaCtg 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1965
               TabIndex        =   46
               ToolTipText     =   "VaCtg Group By Yes/No"
               Top             =   2235
               Width           =   5130
               Begin VB.OptionButton OptVaCtg 
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
                  Left            =   3780
                  TabIndex        =   14
                  ToolTipText     =   "VaCtg Group By Yes/No"
                  Top             =   0
                  Width           =   945
               End
               Begin VB.OptionButton OptVaCtg 
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
                  TabIndex        =   13
                  ToolTipText     =   "VaCtg Group By Yes/No"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.OptionButton OptInvDt 
               Caption         =   "Invoice Date"
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
               Left            =   1965
               TabIndex        =   15
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   2655
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.OptionButton OptInvDt 
               Caption         =   "Print Date"
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
               Left            =   5760
               TabIndex        =   16
               ToolTipText     =   "Select if to show Invoice Date Or Invoice Printing Date"
               Top             =   2655
               Visible         =   0   'False
               Width           =   1965
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   0
               Left            =   2745
               TabIndex        =   2
               ToolTipText     =   "Enter Invoice Year"
               Top             =   1005
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYFR"
               IdName          =   "UOINYYFR"
               CmpStr          =   "InYy> ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   2
               Left            =   3165
               TabIndex        =   3
               ToolTipText     =   "Enter Invoice Character"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRFR"
               IdName          =   "UOINCHRFR"
               CmpStr          =   "InChr >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   3
               Left            =   3945
               TabIndex        =   4
               ToolTipText     =   "Enter Invoice Number"
               Top             =   1005
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOFR"
               IdName          =   "UOINNOFR"
               CmpStr          =   "InNo >="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   17
               Left            =   1965
               TabIndex        =   0
               ToolTipText     =   "Enter Company Code"
               Top             =   715
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "InCoCd = "
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   1
               Left            =   1965
               TabIndex        =   1
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   1005
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCFR"
               IdName          =   "UOINTCFR"
               CmpStr          =   "InTc> ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   4
               Left            =   1965
               TabIndex        =   10
               ToolTipText     =   "Specify Whether To Show Actual, Invoice Or Customs Values"
               Top             =   1575
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOINVAL"
               IdName          =   "UOINVAL"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   5
               Left            =   6540
               TabIndex        =   6
               ToolTipText     =   "Enter Invoice Year"
               Top             =   960
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOINYYTO"
               IdName          =   "UOINYYTO"
               CmpStr          =   "InYy< ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   6
               Left            =   6960
               TabIndex        =   7
               ToolTipText     =   "Enter Invoice Character"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINCHRTO"
               IdName          =   "UOINCHRTO"
               CmpStr          =   "InChr< ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   7
               Left            =   7740
               TabIndex        =   8
               ToolTipText     =   "Enter Invoice Number"
               Top             =   960
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOINNOTO"
               IdName          =   "UOINNOTO"
               CmpStr          =   "InNo< ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   8
               Left            =   5760
               TabIndex        =   5
               ToolTipText     =   "Enter Invoice Tc"
               Top             =   960
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOINTCTO"
               IdName          =   "UOINTCTO"
               CmpStr          =   "InTc< ="
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   10
               Left            =   1965
               TabIndex        =   9
               ToolTipText     =   "Enter Company Code"
               Top             =   1290
               Width           =   9105
               _ExtentX        =   16060
               _ExtentY        =   503
               MaxLength       =   255
               DataType        =   4
               DataField       =   "UOINEXPNOSEL"
               IdName          =   "UOINEXPNOSEL"
               CmpStr          =   "InExpNo In"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   46
               Left            =   5760
               TabIndex        =   12
               ToolTipText     =   "Enter To Invoice Voucher Date"
               Top             =   1920
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   48
               Left            =   1965
               TabIndex        =   11
               ToolTipText     =   "Enter From Invoice Voucher Date"
               Top             =   1860
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
            End
            Begin MwfCtl.MWCTL_MED ATXT 
               Height          =   285
               Index           =   80
               Left            =   8745
               TabIndex        =   42
               ToolTipText     =   "Show invoice Date/ Printng Dat (Yes / No)"
               Top             =   2655
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
               Index           =   9
               Left            =   8760
               TabIndex        =   47
               ToolTipText     =   "VaCtg Group By Yes/No"
               Top             =   2280
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               IdName          =   "WVACTGYN"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Group By VaCtg"
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
               Left            =   165
               TabIndex        =   45
               Top             =   2235
               Width           =   1575
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Inv Dt /Print Dt"
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
               Left            =   165
               TabIndex        =   44
               Top             =   1860
               Width           =   1785
            End
            Begin VB.Label LblConsDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Consider Date"
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
               Left            =   165
               TabIndex        =   43
               Top             =   2655
               Visible         =   0   'False
               Width           =   1695
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
               Index           =   9
               Left            =   2880
               TabIndex        =   41
               Top             =   240
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
               Index           =   8
               Left            =   5745
               TabIndex        =   40
               Top             =   240
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
               Index           =   7
               Left            =   6480
               TabIndex        =   39
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
               Index           =   6
               Left            =   6900
               TabIndex        =   38
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
               Index           =   0
               Left            =   7665
               TabIndex        =   37
               Top             =   960
               Width           =   105
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
               Left            =   165
               TabIndex        =   36
               Top             =   715
               Width           =   1335
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
               Index           =   1
               Left            =   3870
               TabIndex        =   35
               Top             =   1005
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
               Index           =   3
               Left            =   3105
               TabIndex        =   34
               Top             =   1005
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
               Index           =   4
               Left            =   2685
               TabIndex        =   33
               Top             =   1005
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
               Index           =   2
               Left            =   165
               TabIndex        =   32
               Top             =   1005
               Width           =   1815
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
               Left            =   165
               TabIndex        =   31
               Top             =   1290
               Width           =   1425
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
               Left            =   165
               TabIndex        =   30
               Top             =   1575
               Width           =   1755
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   24
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpInvLossSubRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rep As New EmrRepInvLossSubRep
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  ''GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub
Private Sub adc_setreprecsource()
  '****** (07/11/06) replaced IrRmCd by RmKt ******
  '*** Report Sql ***
  '  Select InTc, InYy, InChr, InNo, InExpNo, InCmCd, CmName, InDt, IdGldLs, RmCtg,
  '  RmKt, IrRmIWt as qMetWt, IrRmIWt as qMetPureWt, IrRmIWt as qInvLsWt,
  '  IrRmIWt As qInvPureLsWt, IrRmIWt As qInvTotWt, IrRmIWt As qInvPureTotWt
  '  from InvRm, RmMst, InvDsg, InvHd, CustMst where 1= 2
  '*** Report Sql ***
  
  Dim wAddlCnd As String '****** Sachin 3.02.0 - New fields in Where Clause
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  
  Call DispCoNm
  '**************Geeta******Emr206**05/04/04*******
  If UCase(adc("UoInVal")) = "A" Then
    wIrRmWt = "IrRmAWt"
  ElseIf UCase(adc("UoInVal")) = "S" Then
    wIrRmWt = "IrRmAWt"
  ElseIf UCase(adc("UoInVal")) = "I" Then
    wIrRmWt = "IrRmIWt"
  ElseIf UCase(adc("UoInVal")) = "Z" Then
    wIrRmWt = "IrRmZWt"
  End If
  '********************
  Dim wInvDt As String
  wInvDt = "InDt "
  If UCase(adc.MenuCd) <> UCase("RepEPSPlainJwl") Then
    If adc("UoYN2") = "Y" Then
      wInvDt = "InDt "
    Else
      wInvDt = "InPrnDt "
    End If
      adc("UoDtFr").CmpStr = wInvDt + ">= "
      adc("UoDtTo").CmpStr = wInvDt + "<= "
  End If
      
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  'PREVIOUS - EMR205 <<<<<<====================================
  'wSqlStrg = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
             "max(InNo) as InNo, max(InExpNo), max(InCmCd), " + _
             "(Select CmName from CustMst where CmCtg= 'C' and CmCd= max(InCmCd)) as qCmName, " + _
             "max(InDt), IdGldLs, " + _
             "max(RmCtg) as RmCtg, IrRmCd, sum(IrRmIWt) as qMetWt, " + _
             "sum(IrRmIWt * RmPurityZ) as qMetPureWt, " + _
             "sum(IrRmIWt * IdGldLs/100) as qInvLsWt, " + _
             "sum(IrRmIWt * RmPurityZ * IdGldLs/100) as qInvPureLsWt, " + _
             "sum(IrRmIWt * (1 + IdGldLs/100)) as qInvTotWt, " + _
             "sum(IrRmIWt * RmPurityZ * (1 + IdGldLs/100)) as qInvPureTotWt " + _
             " from InvRm join RmMst on RmCd = IrRmCd " + _
             "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
             "join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy " + _
             "and InChr= IdChr and InNo= IdNo " + _
             "where RmCtg in ('G', 'P') " + _
             wCnd + " Group By IrRmCd, IdGldLs "
  '**************Geeta********Emr206*****05/04/04
  ' ****** Sachin 2.12 - 05/12/05 - SJM - Silver added to the query
  
  '*** (Bef 2.13)
  'wSqlStrg = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
  '           "max(InNo) as InNo, max(InExpNo) as InExpNo, max(InCmCd) as InCmCd, " + _
  '           "(Select CmName from CustMst where CmCtg= 'C' and CmCd= max(InCmCd)) as CmName, " + _
  '           "max(InDt) as InDt, IdGldLs, " + _
  '           "max(RmCtg) as RmCtg, IrRmCd, sum(" + wIrRmWt + ") as qMetWt, " + _
  '           "sum(" + wIrRmWt + " * RmPurityZ) as qMetPureWt, " + _
  '           "sum(" + wIrRmWt + " * IdGldLs/100) as qInvLsWt, " + _
  '           "sum(" + wIrRmWt + "* IdGldLs/100 * RmPurityZ) as qInvPureLsWt, " + _
  '           "sum(" + wIrRmWt + "* (1 + IdGldLs/100)) as qInvTotWt, " + _
  '           "sum(" + wIrRmWt + "* (1 + IdGldLs/100) * RmPurityZ) as qInvPureTotWt " + _
  '           " from InvRm join RmMst on RmCd = IrRmCd " + _
  '           "join InvDsg on IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + _
  '           "join InvHd on InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy " + _
  '           "and InChr= IdChr and InNo= IdNo " + _
  '           "where RmCtg in ('G', 'P', 'S') " + _
  '           wCnd + " Group By IrRmCd, IdGldLs "
  '*** (Bef 2.13)
  
  '****** (Jen 2.14 Next Patch 1 05/06/07) added R1 alias, R2 join and "/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)"
  '*** (Jen 2.13)
  '****** (07/11/06) IrRmCd has been replaced by RmKt in the field list as well as Group By ******
  '****** Sachin 3.02 - Id fields in Joins
  
  ' **** Manali 3.03 - 25/06/08 - Pd Mod - 'L' added
  '578.44 qIdVaCtg,qRmdesc,qVactgdesc and param join for rmctg,vactg added
  'rmctg, vactg added in group by
  wSqlStrg = "Select max(InTc) as InTc, max(InYy) as InYy, max(InChr) as InChr, " + _
             "max(InNo) as InNo, max(InExpNo) as InExpNo, max(InCmCd) as InCmCd, " + _
             "(Select CmName from CustMst where CmCtg= 'C' and CmCd= max(InCmCd)) as CmName, " + _
             "max(InDt) as InDt, IdGldLs, " + _
             "R1.RmCtg as RmCtg, R1.RmKt, sum(" + wIrRmWt + ") as qMetWt, " + _
             "sum(" + wIrRmWt + " * R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qMetPureWt, " + _
             "sum(" + IIF(UCase(adc("UoInVal")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " *(Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100 end) ", wIrRmWt + "* (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100") + ") as qInvLsWt, " + _
             "sum(" + IIF(UCase(adc("UoInVal")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt- " + wIrRmWt + " else " + wIrRmWt + " * (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100 end) ", wIrRmWt + "* (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100") + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureLsWt, " + _
             "sum(" + IIF(UCase(adc("UoInVal")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wIrRmWt + " * (1 + (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100) end) ", wIrRmWt + "* (1 + (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100)") + ") as qInvTotWt, " + _
             "sum(" + IIF(UCase(adc("UoInVal")) = "Z", " (case when IrMainMet= 'Y' then IdCustmMtchWt else " + wIrRmWt + " * (1 + (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100) end) ", wIrRmWt + "* (1 + (Case when IrRmCtg='M' Then 0 Else IdGldLs End )/100)") + "* R1.RmPurityZ/ (case when R2.RmPurityZ> 0 then R2.RmPurityZ else 1 end)) as qInvPureTotWt " + _
             ",(case when '" + adc("wVaCtgYN") + "'='Y' then IdVaCtg else R1.RmCtg end) as qIdVaCtg,max(Param.PDesc) as qRmdesc, max(VaCtg.PDesc) as qVaCtgdesc " + _
             " from InvRm join RmMst R1 on R1.RmCd = IrRmCd " + IIF(gs_Partition = ctCurrPrtn, " and R1.RmPrtKey=IrPrtKey ", "") + " join RmMst R2 on R2.RmCd= R1.RmZCd " + IIF(gs_Partition = ctCurrPrtn, " and R2.RmPrtKey=R1.RmPrtKey ", "") + _
             " join InvDsg on IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo and IdSr= IrSr " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=IrPrtKey ", "") + _
             " join InvHd on InIdNo=IdInIdNo and  InCoCd= IdCoCd and InTc = IdTc and InYy= IdYy " + _
             " and InChr= IdChr and InNo= IdNo " + IIF(gs_Partition = ctCurrPrtn, " and IdPrtKey=InPrtKey ", "") + _
             " join Param on PTyp ='RMCTG' and PMCd =IrRmCtg and PSCd ='' " + _
             " join Param VaCtg on VaCtg.PTyp ='VACTG' and VaCtg.PMCd =IdVaCtg and VaCtg.PSCd ='' " + _
             " where " + IIF(gs_Partition = ctCurrPrtn, " IrPrtKey='" + ctCurrPrtn + "' and ", "") + " R1.RmCtg in ('G', 'P', 'S', 'L','M') And IrRmCtg in ('G', 'P', 'S', 'L','m') " + _
             wCnd + " Group By R1.RmCtg," + IIF(adc("wVaCtgYN") = "Y", "IdVaCtg,", "") + "R1.RmKt, IdGldLs "
  '*** (Jen 2.13)
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 added ctMaxDopOpt
  CRV_REP.DisplayGroupTree = False
  
  Rep.TxtScpInvFr.SetText IIF(adc("UoInTcFr") = "", "  ", adc("UoInTcFr")) + "/" + IIF(adc("UoInYyFr") = "", "  ", adc("UoInYyFr")) + "/" + IIF(adc("UoInChrFr") = "", "  ", adc("UoInChrFr")) + "/" + IIF(adc("UoInNoFr") = 0, "  ", CStr(adc("UoInNoFr")))
  Rep.TxtScpInvTo.SetText IIF(adc("UoInTcTo") = "", "  ", adc("UoInTcTo")) + "/" + IIF(adc("UoInYyTo") = "", "  ", adc("UoInYyTo")) + "/" + IIF(adc("UoInChrTo") = "", "  ", adc("UoInChrTo")) + "/" + IIF(adc("UoInNoTo") = 0, "  ", CStr(adc("UoInNoTo")))
  Rep.TxtScpInvSel.SetText adc("UoInExpNoSel")
  Rep.TxtInvDtFr.SetText adc("UoDtFr")
  Rep.TxtInvDtTo.SetText adc("UoDtTo")
  '578.44 if Vactg=yes then it will display header& footer section of vactg and grand total
  If adc("wVaCtgYN") = "Y" Then
    Rep.GHVaCtg.Suppress = False: Rep.GFVaCtg.Suppress = False
    Rep.TxtGTotal.Suppress = False
    Rep.qMetWt1.Suppress = False
    Rep.qMetPureWt1.Suppress = False: Rep.qInvLsWt1.Suppress = False
    Rep.qInvPureLsWt1.Suppress = False: Rep.qInvTotWt1.Suppress = False
    Rep.qInvPureTotWt1.Suppress = False:  Rep.LnTot.Suppress = False
    Rep.GHRmCtg.Suppress = False: Rep.wTxtFRmCtg1.Suppress = False
    Rep.LnRmCtg.Suppress = True
  End If
  
Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
  
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
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    Set adc.FirNKeyCtl = adc("UoInTcFr") 'adc("UoCoCdFr")
    Call InitProp(Me)
    
    gs_InTcTyp = "IN"
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  gs_InTcTyp = "IN"
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  gs_InTcTyp = "IN"
  SetProp Me, IdName, When
End Sub
'578.44 store Vactg YN value
Private Sub OptVaCtg_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("wVaCtgYN") = "Y"
Case Is = 1
  adc("wVaCtgYN") = "N"
End Select
End Sub
Private Sub OptVaCtg_GotFocus(Index As Integer)
  DispMsg FraVaCtg.ToolTipText, etInfo
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
 adc("wVaCtgYN") = "N"
 If adc("wVaCtgYN") = "Y" Then
  OptVaCtg(0).Value = True
 Else
  OptVaCtg(1).Value = True
 End If

End Sub
Private Sub OptInvDt_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN2") = "Y"
Case Is = 1
  adc("UoYN2") = "N"
End Select
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
