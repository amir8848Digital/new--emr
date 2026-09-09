VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpCertLst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate Listing"
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
      Left            =   10830
      TabIndex        =   0
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
      TabIndex        =   23
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
      Left            =   4785
      TabIndex        =   31
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   32
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
         TabIndex        =   33
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
      Left            =   6000
      TabIndex        =   30
      Top             =   9645
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   28
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
      TabIndex        =   24
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
      TabIndex        =   27
      Top             =   60
      Width           =   14895
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   60
         TabIndex        =   26
         Top             =   30
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpCertLst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpCertLst.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   8715
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   14520
            Begin VB.CheckBox ChkBoxArr 
               Alignment       =   1  'Right Justify
               Caption         =   "Show 2nd Line"
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
               TabIndex        =   21
               Tag             =   "ShowDemarc"
               ToolTipText     =   "Check To Show 2nd Line"
               Top             =   3280
               Width           =   1800
            End
            Begin VB.Frame Fra_RmCdQlty 
               BorderStyle     =   0  'None
               Height          =   640
               Left            =   0
               TabIndex        =   45
               Top             =   2690
               Width           =   14520
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   7
                  Left            =   6495
                  TabIndex        =   18
                  ToolTipText     =   "Enter To Raw Material Code"
                  Top             =   0
                  Width           =   2475
                  _ExtentX        =   4366
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataType        =   4
                  DataField       =   "UORMCDTO"
                  IdName          =   "UORMCDTO"
                  CmpStr          =   "CdRmCd <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   8
                  Left            =   1590
                  TabIndex        =   17
                  ToolTipText     =   "Enter From Raw Material Code "
                  Top             =   0
                  Width           =   2475
                  _ExtentX        =   4366
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataType        =   4
                  DataField       =   "UORMCDFR"
                  IdName          =   "UORMCDFR"
                  CmpStr          =   "CdRmCd >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   15
                  Left            =   1590
                  TabIndex        =   19
                  ToolTipText     =   "Enter From FMBarCode "
                  Top             =   300
                  Width           =   4545
                  _ExtentX        =   8017
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataType        =   4
                  DataField       =   "UOCDFMBARCDFR"
                  IdName          =   "UOCDFMBARCDFR"
                  CmpStr          =   "CdFMBarCd >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   10
                  Left            =   6495
                  TabIndex        =   20
                  ToolTipText     =   "Enter To FMBarcode "
                  Top             =   300
                  Width           =   4545
                  _ExtentX        =   8017
                  _ExtentY        =   503
                  MaxLength       =   30
                  DataType        =   4
                  DataField       =   "UOCDFMBARCDTO"
                  IdName          =   "UOCDFMBARCDTO"
                  CmpStr          =   "CdFMBarCd <="
               End
               Begin VB.Label LblFMBarCd 
                  BackStyle       =   0  'Transparent
                  Caption         =   "FM BarCd"
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
                  TabIndex        =   47
                  Top             =   300
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
                  Left            =   60
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin VB.Frame Fra_Lot 
               BorderStyle     =   0  'None
               Height          =   320
               Left            =   0
               TabIndex        =   41
               Top             =   2080
               Width           =   14520
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   66
                  Left            =   1590
                  TabIndex        =   7
                  ToolTipText     =   "Enter From Lot Number"
                  Top             =   0
                  Width           =   2500
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataType        =   4
                  DataField       =   "UOLOTNOFR"
                  IdName          =   "UOLOTNOFR"
                  CmpStr          =   "CdRmLotNo >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   67
                  Left            =   6500
                  TabIndex        =   8
                  ToolTipText     =   "Enter To Lot Number"
                  Top             =   0
                  Width           =   2500
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  MaxLength       =   16
                  DataType        =   4
                  DataField       =   "UOLOTNOTO"
                  IdName          =   "UOLOTNOTO"
                  CmpStr          =   "CdRmLotNo <="
               End
               Begin VB.Label LblLotNo 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rm Lot No"
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
                  TabIndex        =   42
                  Top             =   0
                  Width           =   1425
               End
            End
            Begin VB.Frame Fra_Bag 
               BorderStyle     =   0  'None
               Height          =   320
               Left            =   0
               TabIndex        =   40
               Top             =   2380
               Width           =   12135
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   47
                  Left            =   1590
                  TabIndex        =   9
                  ToolTipText     =   "Enter From Bag Company Code"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOCOCDFR"
                  IdName          =   "UOCOCDFR"
                  CmpStr          =   "CdBCoCd >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   36
                  Left            =   2400
                  TabIndex        =   10
                  ToolTipText     =   "Enter From Bag Year"
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UOBYYFR"
                  IdName          =   "UOBYYFR"
                  CmpStr          =   "CdBYy >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   38
                  Left            =   2840
                  TabIndex        =   11
                  ToolTipText     =   "Enter From Bag Character"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOBCHRFR"
                  IdName          =   "UOBCHRFR"
                  CmpStr          =   "CdBChr >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   41
                  Left            =   3630
                  TabIndex        =   12
                  ToolTipText     =   "Enter From Bag Number"
                  Top             =   0
                  Width           =   925
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOBNOFR"
                  IdName          =   "UOBNOFR"
                  CmpStr          =   "CdBNo >="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   0
                  Left            =   6500
                  TabIndex        =   13
                  ToolTipText     =   "Enter To Company Code"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOCOCDTO"
                  IdName          =   "UOCOCDTO"
                  CmpStr          =   "CdBCoCd <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   1
                  Left            =   7290
                  TabIndex        =   14
                  ToolTipText     =   "Enter To Bag Year"
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   503
                  MaxLength       =   2
                  DataField       =   "UOBYYTO"
                  IdName          =   "UOBYYTO"
                  CmpStr          =   "CdBYy <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   2
                  Left            =   7750
                  TabIndex        =   15
                  ToolTipText     =   "Enter To Bag Character"
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  MaxLength       =   3
                  DataType        =   4
                  DataField       =   "UOBCHRTO"
                  IdName          =   "UOBCHRTO"
                  CmpStr          =   "CdBChr <="
               End
               Begin MwfCtl.MWCTL_MED atxt 
                  Height          =   285
                  Index           =   3
                  Left            =   8560
                  TabIndex        =   16
                  ToolTipText     =   "Enter To Bag Number"
                  Top             =   0
                  Width           =   925
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  MaxLength       =   0
                  DataType        =   1
                  DataField       =   "UOBNOTO"
                  IdName          =   "UOBNOTO"
                  CmpStr          =   "CdBNo <="
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
                  Left            =   8480
                  TabIndex        =   54
                  Top             =   0
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
                  Index           =   5
                  Left            =   7680
                  TabIndex        =   53
                  Top             =   0
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
                  Left            =   7220
                  TabIndex        =   52
                  Top             =   0
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
                  Left            =   3560
                  TabIndex        =   51
                  Top             =   0
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
                  Left            =   2760
                  TabIndex        =   50
                  Top             =   0
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
                  Left            =   2330
                  TabIndex        =   49
                  Top             =   0
                  Width           =   105
               End
               Begin VB.Label LblBagYyChr 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bag No"
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
                  Top             =   0
                  Width           =   1245
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   42
               Left            =   1590
               TabIndex        =   1
               ToolTipText     =   "Enter Lot Or Bag Or HUID"
               Top             =   930
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOCDCTG"
               IdName          =   "UOCDCTG"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1590
               TabIndex        =   5
               ToolTipText     =   "Enter From Certificate \ HUID Number"
               Top             =   1785
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   503
               MaxLength       =   30
               DataType        =   4
               DataField       =   "UOCDCERTNOFR"
               IdName          =   "UOCDCERTNOFR"
               CmpStr          =   "CdCertNo >= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   6495
               TabIndex        =   6
               ToolTipText     =   "Enter To Certificate \ HUID Number"
               Top             =   1785
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   503
               MaxLength       =   30
               DataType        =   4
               DataField       =   "UOCDCERTNOTO"
               IdName          =   "UOCDCERTNOTO"
               CmpStr          =   "CdCertNo <= "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   28
               Left            =   6500
               TabIndex        =   4
               ToolTipText     =   "Enter To Certificate \ HUID Date"
               Top             =   1505
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTTO"
               IdName          =   "UODTTO"
               CmpStr          =   "CdDt <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   30
               Left            =   1590
               TabIndex        =   3
               ToolTipText     =   "Enter From Certificate \ HUID Date"
               Top             =   1505
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UODTFR"
               IdName          =   "UODTFR"
               CmpStr          =   "CdDt >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1590
               TabIndex        =   2
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   1215
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "CdRmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   1900
               TabIndex        =   55
               ToolTipText     =   "Show Demarcation Line(Y/ N)"
               Top             =   3240
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   6
               Left            =   1590
               TabIndex        =   22
               ToolTipText     =   "Enter Level Of Detail To Be Shown"
               Top             =   3560
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   4
               DataField       =   "UOPMCDFR"
               IdName          =   "UOPMCDFR"
            End
            Begin VB.Label LblSortBy 
               BackStyle       =   0  'Transparent
               Caption         =   "Sort By"
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
               TabIndex        =   48
               Top             =   3560
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
               Left            =   60
               TabIndex        =   44
               Top             =   1215
               Width           =   1425
            End
            Begin VB.Label LblLotOrBag 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot / Bag / HUID"
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
               TabIndex        =   39
               Top             =   930
               Width           =   1545
            End
            Begin VB.Label LblDt 
               BackStyle       =   0  'Transparent
               Caption         =   "Cert Date"
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
               TabIndex        =   38
               Top             =   1505
               Width           =   1425
            End
            Begin VB.Label LblCertNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Cert No"
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
               TabIndex        =   37
               Top             =   1790
               Width           =   1425
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
               Left            =   1590
               TabIndex        =   36
               Top             =   480
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
               Left            =   6500
               TabIndex        =   35
               Top             =   480
               Width           =   1275
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   -74940
            TabIndex        =   29
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   25
               Top             =   360
               Width           =   8805
               _ExtentX        =   15531
               _ExtentY        =   635
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpCertLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepCertLst
Dim moCn As MwfLib.MDOConnection

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
 If adc("UoYN") = "" Then adc("UoYN") = "N"
 If adc("UoYN") = "Y" Then ChkBoxArr(0).Value = Checked Else ChkBoxArr(0).Value = Unchecked
 '62 HUID heading updated
  If adc("UoCdCtg") = "H" Then
    LblDt.Caption = "HUID Date"
    LblCertNo.Caption = "HUID No"
  Else
    LblDt.Caption = "Cert Date"
    LblCertNo.Caption = "Cert No"
  End If
  
End Sub

Private Sub adc_setreprecsource()
On Error GoTo RepErr
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, ws_OrdBy As String
  Dim ws_DetOrdBy As String
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)

  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")
  

  If adc("UoPMCdFr") = "CDBAGNO" Then
    ws_DetOrdBy = " ,CdRmCtg, CdRmLotNo,CdCertNo"
  ElseIf adc("UoPMCdFr") = "CDCERT" Then
    ws_DetOrdBy = " ,CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo"
  ElseIf adc("UoPMCdFr") = "CDDT" Then
    ws_DetOrdBy = " ,CdRmCtg , CdRmLotNo , CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
  ElseIf adc("UoPMCdFr") = "CDRMLOT" Then
    ws_DetOrdBy = " , CdBCoCd, CdBYy ,CdBChr, CdBNo, CdCertNo"
  End If
  
  If adc("UoPMCdFr") <> "" Then
    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp='SRTCD' and PMCd='" + adc("UoPMCdFr") + "'") + ws_DetOrdBy
  ElseIf adc("UoPMCdFr") = "" Then
    ws_OrdBy = "Order By CdRmCtg,CdRmLotNo,CdBCoCd, CdBYy ,CdBChr, CdBNo,CdCertNo"
  End If
   
  
  If adc("UoCdCtg") = "L" Then
  '11.b CdSubShp added
   wSqlStrg = "Select  " + _
             "CdDt,CdRmLotNo,CdCertNo,CdRmCtg, CdRmQty,CdRmCol,CdRmQlty, " + _
             " CdRmCd,CdRmWt,CdFMBarCd,CdComment,CdDocNo,CdDocDt, CdRem,  " + _
             " '' as CdBCoCd, '' as CdBYy, '' as CdBChr, 0 as CdBNo, CdSubShp " + _
             "from CertDet where CdCtg='L' " + wCnd + ws_OrdBy
  '62 ADC("UoCdCtg") = "H" added
  '11.b CdSubShp added
  '99.f sub shape not required for B & H
  ElseIf adc("UoCdCtg") = "B" Or adc("UoCdCtg") = "H" Then
 
  wSqlStrg = "Select  " + _
             "CdDt,CdBCoCd, CdBYy, CdBChr, CdBNo,CdCertNo, CdComment, CdDocNo,CdDocDt,CdRem, " + _
             " '' as CdRmLotNo, '' as CdRmCtg, 0 as CdRmQty, '' as CdRmCol,  '' as CdRmQlty, " + _
             "'' as CdRmCd, 0 as CdRmWt, '' as CdFMBarCd, '' as CdSubShp " + _
             "from CertDet where CdCtg='" + adc("UoCdCtg") + "'" + wCnd + ws_OrdBy
             
  End If
 
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt
  
  If adc("UoCdCtg") = "L" Then
    Rep.TxtHead.SetText "Certificate Listing - Rm Lot"
    Rep.Hd2.Suppress = True
    Rep.DET2.Suppress = True
    Rep.DET2A.Suppress = True
        
     If adc("UoYN") = "N" Then
        Rep.TxtCmnt1.Suppress = True
        Rep.TxtDocNo1.Suppress = True
        Rep.TxtDocDt1.Suppress = True
        Rep.TxtRem1.Suppress = True
        Rep.hSubShp.Suppress = True
        Rep.DET1A.Suppress = True
        Rep.DET2B.Suppress = True
     End If
  '62 ADC("UoCdCtg") = "H" added and HUID heading updated
  ElseIf adc("UoCdCtg") = "B" Or adc("UoCdCtg") = "H" Then
    If adc("UoCdCtg") = "H" Then
      Rep.TxtHead.SetText "HUID Listing "
      Rep.hCertNo.SetText "HUID No"
      Rep.hCertDt.SetText "HUID Dt"
      Rep.FrCertDt.SetText "Fr HUID Dt "
      Rep.ToCertDt.SetText "To HUID Dt "
      Rep.FrCertNo.SetText "Fr HUID No "
      Rep.ToCertNo.SetText "To HUID No "
    Else
      Rep.TxtHead.SetText "Certificate Listing - Bag"
      Rep.hCertNo.SetText "Certificate No"
      Rep.hCertDt.SetText "Cert Dt"
      Rep.FrCertDt.SetText "Fr Cert Dt "
      Rep.ToCertDt.SetText "To Cert Dt "
      Rep.FrCertNo.SetText "Fr Cert No "
      Rep.ToCertNo.SetText "To Cert No "
    End If
    Rep.Hd1.Suppress = True
    Rep.DET1.Suppress = True
    Rep.DET1A.Suppress = True
    
    If adc("UoYN") = "N" Then
        Rep.TxtCmnt2.Suppress = True
        Rep.TxtRem2.Suppress = True
        Rep.DET2A.Suppress = True
        Rep.DET2B.Suppress = True
    End If
    
  End If
  
  If adc("UoCoCdFr") = "" And adc("UoBYyFr") = "" And adc("UoBChrFr") = "" And adc("UoBNoFr") = 0 And _
      adc("UoCoCdTo") = "" And adc("UoBYyTo") = "" And adc("UoBChrTo") = "" And adc("UoBNoTo") = 0 Then
    Rep.Slash1.Suppress = True: Rep.Slash2.Suppress = True: Rep.Slash3.Suppress = True: Rep.TxtFrBNo.Suppress = True
    Rep.Slash4.Suppress = True: Rep.Slash5.Suppress = True: Rep.Slash6.Suppress = True: Rep.TxtToBNo.Suppress = True
        
  End If
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.CdRmCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  Call DispReport(adc)
  
Exit Sub
RepErr:
  Call ChkSelScope(Err.Number, Err.Description)
End Sub

Private Sub ChkBoxArr_Click(Index As Integer)
 With ChkBoxArr(Index)
 Select Case Index
   Case Is = 0    'Tag = "ShowLine"
    If .Value = Checked Then adc("UoYN") = "Y" Else adc("UoYN") = "N"
  End Select
 End With
End Sub
Private Sub ChkBoxArr_GotFocus(Index As Integer)
  DispMsg ChkBoxArr(Index).ToolTipText, etInfo
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
    Set adc.FirNKeyCtl = adc("UoCdCtg")
    Call InitProp(Me)
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
  '62 HUID heading updated
  Select Case UCase(IdName)
    Case Is = UCase("UoCdCtg")
      If pv_NewValue = "H" Then
        LblDt.Caption = "HUID Date"
        LblCertNo.Caption = "HUID No"
      Else
        LblDt.Caption = "Cert Date"
        LblCertNo.Caption = "Cert No"
      End If
  End Select
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  Select Case UCase(IdName)
    Case Is = UCase("UoLotNoFr"), UCase("UoLotNoTo")
      If adc.Mode = xNorm Then
        HlpList.RmLotNo (adc("UoRmCtgFr")): Exit Sub
      End If
      Case Is = UCase("UoPMCdFr")
      If adc.Mode = xNorm Then
        HlpList.PMCd "SRTCD": Exit Sub
      End If
      Case Is = UCase("UoRmCdFr"), UCase("UoRmCdTo")
      If adc.Mode = xNorm Then
      HlpList.RmCd "N", "'" + adc("UoRmCtgFr") + "'": Exit Sub
      End If
      
  End Select
  SetProp Me, IdName, When
End Sub

Private Sub tab_REP_Click(PreviousTab As Integer)
    fra_tabrep(TAB_REP.Tab).Enabled = True
    fra_tabrep(PreviousTab).Enabled = False
End Sub
Private Sub ADC_KeyWhen()
    adc("UoUsrCd") = adc.UsrCd
    adc("UoMnuCd") = adc.MenuCd
   
End Sub
Private Sub DispCoNm()
  Rep.wCoCd.SetText gs_CoNm

  Rep.wCoCdLogo.SetText gs_CoNm
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If

  Dim ws_Dt As String, ws_Tm As String
    ws_Dt = Format$(moCn.SrvrDate, IIF(UCase(MWLib.DtSetting) = "MM/DD/YY", "MMM-DD-YYYY", "DD-MMM-YYYY"))
    ws_Tm = Replace(Format(IIF(CDbl(moCn.SrvrTime) >= 13#, moCn.SrvrTime - 12, IIF(CDbl(moCn.SrvrTime) < 1#, moCn.SrvrTime + 12, moCn.SrvrTime)), "00.00"), ".", ":") + IIF(CDbl(moCn.SrvrTime) >= 12#, " PM", " AM")
  Rep.TxtRepUsrDtTime.SetText Trim(adc.UsrCd) + " / " + ws_Dt + " / " + ws_Tm

  Rep.TxtCdCtg.SetText adc("UoCdCtg"):  Rep.TxtRmCtg.SetText adc("UoRmCtgFr")
  Rep.TxtFrCertDt.SetText adc("UoDtFr"):   Rep.TxtToCertDt.SetText adc("UoDtTo")
  Rep.TxtFrCertNo.SetText adc("UoCdCertNoFr"):  Rep.TxtToCertNo.SetText adc("UoCdCertNoTo")
  Rep.TxtFrRmLotNo.SetText adc("UoLotNoFr"):  Rep.TxtToRmLotNo.SetText adc("UoLotNoTo")
  Rep.TxtFrBCoCd.SetText adc("UoCoCdFr"):  Rep.TxtFrBYy.SetText adc("UoBYyFr")
  Rep.TxtFrBChr.SetText adc("UoBChrFr"):  Rep.TxtFrBNo.SetText adc("UoBNoFr")
  Rep.TxtToBCoCd.SetText adc("UoCoCdTo"):  Rep.TxtToBYy.SetText adc("UoBYyTo")
  Rep.TxtToBChr.SetText adc("UoBChrTo"):  Rep.TxtToBNo.SetText adc("UoBNoTo")
  Rep.TxtFrRmCode.SetText adc("UoRmCdFr"):  Rep.TxtToRmCode.SetText adc("UoRmCdTo")
  Rep.TxtFrFMBarCd.SetText adc("UoCdFMBarCdFr"):  Rep.TxtToFMBarCd.SetText adc("UoCdFMBarCdTo")

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


