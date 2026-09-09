VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpBagReco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bag Reconciliation"
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
      Left            =   10980
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9660
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MwfCtl.MWCTL_DST DST_REP 
      Height          =   570
      Left            =   4440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9630
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1005
   End
   Begin VB.Frame fra_keys 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4905
      TabIndex        =   26
      Top             =   9630
      Width           =   3555
      Begin MwfCtl.MWCTL_MED aktxt 
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   27
         Top             =   105
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
         Top             =   120
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
      Left            =   6120
      TabIndex        =   25
      Top             =   9630
      Visible         =   0   'False
      Width           =   855
      Begin CRVIEWER9LibCtl.CRViewer9 CRV_REP 
         Height          =   495
         Left            =   480
         TabIndex        =   22
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
      TabIndex        =   16
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
      Height          =   9615
      Left            =   0
      TabIndex        =   21
      Top             =   60
      Width           =   15360
      Begin TabDlg.SSTab TAB_REP 
         Height          =   9465
         Left            =   15
         TabIndex        =   20
         Top             =   45
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "&Data Range"
         TabPicture(0)   =   "EmrFrpBagReco.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpBagReco.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fra_tabrep(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   9070
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   360
            Width           =   15105
            Begin MwfCtl.MWCTL_GRP GRP_REP 
               Height          =   360
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   7965
               _ExtentX        =   14049
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   -74940
            TabIndex        =   23
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptReco 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1710
               TabIndex        =   44
               ToolTipText     =   "Select If to Show All Or Unmatched Records"
               Top             =   2650
               Width           =   2925
               Begin VB.OptionButton OptReco 
                  Caption         =   "Unmatched"
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
                  TabIndex        =   13
                  ToolTipText     =   "Select If to Show All Or Unmatched Records"
                  Top             =   0
                  Width           =   1485
               End
               Begin VB.OptionButton OptReco 
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
                  TabIndex        =   12
                  ToolTipText     =   "Select If to Show All Or Unmatched Records"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   18
               Left            =   5055
               TabIndex        =   6
               ToolTipText     =   "Enter To Voucher Year"
               Top             =   1410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOTYYTO"
               IdName          =   "UOTYYTO"
               CmpStr          =   "BmYy <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   20
               Left            =   2520
               TabIndex        =   3
               ToolTipText     =   "Enter From Voucher Year"
               Top             =   1410
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               MaxLength       =   2
               DataType        =   4
               DataField       =   "UOTYYFR"
               IdName          =   "UOTYYFR"
               CmpStr          =   "BmYy >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   43
               Left            =   1725
               TabIndex        =   2
               ToolTipText     =   "Enter From Voucher Tc"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTTCFR"
               IdName          =   "UOTTCFR"
               CmpStr          =   "BmTc >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   45
               Left            =   2925
               TabIndex        =   4
               ToolTipText     =   "Enter From Voucher Character"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRFR"
               IdName          =   "UOTCHRFR"
               CmpStr          =   "BmChr >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   46
               Left            =   4275
               TabIndex        =   5
               ToolTipText     =   "Enter To Voucher Tc"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTTCTO"
               IdName          =   "UOTTCTO"
               CmpStr          =   "BmTc <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   47
               Left            =   5475
               TabIndex        =   7
               ToolTipText     =   "Enter To Voucher Character"
               Top             =   1410
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOTCHRTO"
               IdName          =   "UOTCHRTO"
               CmpStr          =   "BmChr <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   29
               Left            =   4275
               TabIndex        =   9
               ToolTipText     =   "Enter To Voucher Number"
               Top             =   1695
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBRNOTO"
               IdName          =   "UOBRNOTO"
               CmpStr          =   "BmNo <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   12
               Left            =   1725
               TabIndex        =   1
               ToolTipText     =   "Enter Company Code"
               Top             =   1125
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               Enabled         =   0   'False
               MaxLength       =   3
               DataType        =   4
               DataField       =   "UOCOCDFR"
               IdName          =   "UOCOCDFR"
               CmpStr          =   "BmCoCd = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   11
               Left            =   1725
               TabIndex        =   10
               ToolTipText     =   "Enter As On Date"
               Top             =   2265
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               MaxLength       =   8
               DataType        =   3
               DataField       =   "UOASOFDT"
               IdName          =   "UOASOFDT"
               CmpStr          =   "BmDt="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   9
               Left            =   4725
               TabIndex        =   11
               ToolTipText     =   "Show All (Y) or UnMatched Records Only (N)"
               Top             =   2650
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
               Index           =   26
               Left            =   1725
               TabIndex        =   8
               ToolTipText     =   "Enter From Voucher Number"
               Top             =   1695
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   503
               MaxLength       =   0
               DataType        =   1
               DataField       =   "UOBRNOFR"
               IdName          =   "UOBRNOFR"
               CmpStr          =   "BmNo >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   35
               Left            =   1725
               TabIndex        =   14
               ToolTipText     =   "Specify Bag Type or ' ' for All"
               Top             =   3060
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOBAGTYP"
               ReCalcParent    =   "UOBCHRFR"
               IdName          =   "UOBAGTYP"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   38
               Left            =   4275
               TabIndex        =   15
               ToolTipText     =   "Enter Bag Character"
               Top             =   3060
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   503
               MaxLength       =   3
               DataField       =   "UOBCHRFR"
               ReCalcOn        =   "UOBAGTYP"
               IdName          =   "UOBCHRFR"
               CmpStr          =   "BChr = "
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   13
               Left            =   1725
               TabIndex        =   0
               ToolTipText     =   "Specify Whether Report to be Taken From Current Or Previous Partition"
               Top             =   840
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   503
               DataType        =   4
               DataField       =   "UOPRTKEYFR"
               IdName          =   "UOPRTKEYFR"
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Partition"
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
               TabIndex        =   43
               ToolTipText     =   "Location"
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label LblBagTyp 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Type"
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
               TabIndex        =   42
               ToolTipText     =   "Location"
               Top             =   3060
               Width           =   1035
            End
            Begin VB.Label LblUoBChrFr 
               BackStyle       =   0  'Transparent
               Caption         =   "Bag Chr"
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
               Left            =   3300
               TabIndex        =   41
               Top             =   3060
               Width           =   885
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Reco Date"
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
               TabIndex        =   40
               Top             =   2265
               Width           =   1425
            End
            Begin VB.Label ALBL 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "All / Diff Records"
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
               Index           =   32
               Left            =   90
               TabIndex        =   39
               Top             =   2650
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
               Index           =   11
               Left            =   90
               TabIndex        =   38
               Top             =   1125
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
               Index           =   27
               Left            =   2445
               TabIndex        =   37
               Top             =   1410
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
               Left            =   4995
               TabIndex        =   36
               Top             =   1410
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
               Left            =   2865
               TabIndex        =   35
               Top             =   1410
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
               Index           =   24
               Left            =   5415
               TabIndex        =   34
               Top             =   1410
               Width           =   105
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch No/Sr"
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
               Left            =   90
               TabIndex        =   33
               Top             =   1695
               Width           =   1485
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Vch Tc/Yy/Chr"
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
               TabIndex        =   32
               Top             =   1410
               Width           =   1785
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
               TabIndex        =   31
               Top             =   480
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
               TabIndex        =   30
               Top             =   480
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
               Left            =   1725
               TabIndex        =   29
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpBagReco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- General Comments on the Report --------'
' This report gives all the details of a bag in a particular format. The output of this
' report is used as a sticker on the Bags. This report helps the Worker in knowing how much
' raw materials of a particular Rm Ctg has to be put in the bag. The report also guides the
' movement of the bags through the process locations as it specifies the sequence of
' locations. The report helps in identifying a bag (i.e. to which order or customer it belongs)
' It also gives specific instructions for production to the workers regarding the bag
' This report does not have any group options
' The report has a sub report which gives a Rm Category wise summary of the Rm Qty & Wt
'--------------------------------------------'

Option Explicit
Dim Rep As New EmrRepBagReco
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06


Private Sub SetGroupSort()

  GRP_REP.Add "Voucher", "BmTc+'/'+BmYy+'/'+BmChr+'/'+LTrim(str(BmNo))", "BmTc+'/'+BmYy+'/'+BmChr+'/'+LTrim(Str(BmNo))", "wVch", "hVchNo", "", "", "", "'(Co:'+BmBagCoCdFr+'-'+BmBagCoCdTo+'-'+LTrim(BmBagCoCdSel)+') (Loc:'+BmBagLocFr+'-'+BmBagLocTo+'-'+LTrim(BmBagLocSel)+')'"
  
  'pg.19 - Order No., Customer code, Bag typ & Bag chr group sorts are not applicable in Stock Reco
  If UCase(adc.MenuCd) <> UCase("RepJTStkReco") Then
    GRP_REP.Add "Order No", "BOdTc+BOdYy+BOdChr+str(BOdNo)", "BOdTc+'/'+BOdYy+'/'+BOdChr+'/'+LTrim(Str(BOdNo))", "wOrdNo,OmCmCd", "hOrdNo,hOmCmCd", "Customer Code", "", "", ""
    GRP_REP.Add "Customer Code", "OmCmCd", "", "OmCmCd", "hOmCmCd", "", "", "", "(Select CmName From CustMst Where CmCtg= 'C' and CmCd= OmCmCd) "
    
    '*** Jay 2.13(CT) ***
    'Uma *** blank <> 'D'
    GRP_REP.Add "Bag Type", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) ", _
                            "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) ", "", "", "Bag Typ+Chr", "", "", _
                            "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) "
    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
    GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) +BChr", "(case PValue when 'F' then 'Flute Bag' when 'P' then 'Primary Bag' when 'C' then 'Component Bag' when 'D' then 'PD Bag' end) + '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
  End If
' Emr 2.11.0 Correction
'  'If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + gs_CoCd + "' and HCd= '" + ctSelfCmCd + "' ") = "Y" Then
'    GRP_REP.Add "Bag Type", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end )", "", "", "Bag Typ+Chr", "", "", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)"
'    GRP_REP.Add "Bag Character", "BChr", "BChr", "", "", "Bag Type,Bag Typ+Chr", "", "", "(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'    GRP_REP.Add "Bag Typ+Chr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+BChr", "(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end) + '/ ' + BChr", "", "", "Bag Type,Bag Character", "", "", "'('+(case PValue when 'Y' then 'Flute Bag' when 'N' then 'Primary Bag' end)+') '+(Select PDesc from param where PTyp ='BCHR' and PMCd= BChr and PSCd= '')"
'  'End If
'  ' Emr 2.11.0 Correction
'*** Jay 2.13(CT) ***
End Sub

Private Sub adc_setreprecsource()
'*** Report Sql (Used in the Report designer) ***
'    Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
'    space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
'    space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
'    BmTc, BmYy, BmChr, BmNo, BdSr, BmLoc, BmDt, BdBYy, BdBChr, BdBNo, BdDmCd,
'    BdGrWt , BdBQty, BdExtQty, BdUnAcQty, BdOthLocBg, BmRecoYn, OmCmCd
'    from BgRecoMst, BgRecoDet, Bag, OrdMst where 1= 2
'*** Report Sql
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wUnMatched As String
  
  '*** Sachin 2.11.0 ***
    Dim ws_BagJoin As String
  '*** Sachin 2.11.0 ***
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
       grpflds = GRP_REP.GrpFldLst
  
  'Ignore records that have no extra or unmatched quantities
  If adc("UoYN") = "N" Then
    wUnMatched = " ((BdExtQty <> 0) Or (BdUnAcQty <> 0))  And "
  End If
  
'*** Jay 2.13(CT) ***
'Uma *** blank <> 'D'
  ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                    IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                    IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    ' **** Sachin 211 (From WIP Bag Pcs By Zubin)****
'    If UCase(adc("UoBagTyp")) = "N" Then
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'N' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    ElseIf UCase(adc("UoBagTyp")) = "Y" Then
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' and PValue = 'Y' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    Else
'      ws_BagJoin = " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
'              IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ")
'    End If
'  ' **** Sachin 211 **** '
'*** Jay 2.13(CT) ***
  
  'Check for blank date
  If adc("UOASOFDT") = "01/01/80" Then
    adc.RepRecSource = "Select BmTc, BmYy, BmChr, BmNo, BdSr, BmBagLocFr, BmBagLocTo, BmBagLocSel, BmBagCoCdFr, BmBagCocdTo, BmBagCoCdSel, BmDt, BdBYy, BdBChr, BdBNo, BdDmCd, " + _
                       "BdGrWt, BdActualQty, BdBQty, BdExtQty, BdUnAcQty, BdOthLocBg, BdOthLoc, BmRecoYn, OmCmCd " + _
                       " From BgRecoMst,BgRecoDet,OrdMst Where 1=2" + ctMaxDopOpt '****** Sachin 3.01 ctMaxDopOpt added
    Exit Sub
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " Where BmRecoYn='Y' And " + wUnMatched + wrepcnd, "")
  
  'pg.19 - Bag & OrdMst joins are not applicable for Stk Reco
  Dim ws_Join As String
  If UCase(adc.MenuCd) <> UCase("RepJTStkReco") Then
    ws_Join = " Join Bag On BCoCd= BdBagCoCd And BYy=BdBYy And BChr=BdBChr And BNo=BdBNo " + IIF(gs_Partition = ctCurrPrtn, " and BPrtkey='" + ctCurrPrtn + "' ", "") + _
              " Join OrdMst On OmCoCd=BCoCd And OmTc=BOdTc And OmYy=BOdYy And OmChr=BOdChr And OmNo=BOdNo " + IIF(gs_Partition = ctCurrPrtn, " And OmPrtKey=BPrtKey ", "") + _
              " Join Param On PTyp= 'BCHR' and PMCd= BChr and PSCd= '' " + _
                IIF(adc("UoBagTyp") <> "", " and PValue= '" + adc("UoBagTyp") + "' ", " and PValue <> 'D'") + _
                IIF(adc("UOBCHRFR") <> "", " and PMCd= '" + adc("UOBCHRFR") + "' ", " ") + _
              " Join DsgMst On DmIdNo=BDmIdNo and DmCd=BOdDmCd And DmTcTyp='DM' And DmSz='' " + IIF(gs_Partition = ctCurrPrtn, " and DmPrtKey=BPrtKey ", "")
  Else
    ws_Join = " Join DsgMst On DmCd=BdDmCd And DmTcTyp='DM' And DmSz='' "
  End If
  
  '******Geeta*****************Emr206****05/04/04*******
  '**********If UserLevel>6 then don't display the Customer Name
  'Replaced OdDmPrdInst,OdCmPrdInst  with Replace( ,char(13)+ char(10),' ')
  '****** Sachin 3.02 - Id fields in Joins
  wSqlStrg = " Select " + grpflds + ", " + _
             " BmCoCd, BmTc, BmYy, BmChr, BmNo, BdSr, BmBagLocFr, BmBagLocTo, BmBagLocSel, BmBagCoCdFr, BmBagCocdTo, BmBagCoCdSel, BmDt, BdBagCoCd,BdBYy, BdBChr, BdBNo, BdDmCd," + _
             " BdGrWt , (Case When BdActualQty < 0 Then 0 Else BdActualQty End) As BdActualQty, BdBQty, BdExtQty, BdUnAcQty, BdOthLocBg, BdOthLoc, BmRecoYn " + _
               IIF(UCase(adc.MenuCd) <> UCase("RepJTStkReco"), ", OmCmCd", "") + _
             " From BgRecoMst " + _
             " Join BgRecoDet On BdCoCd=BmCoCd And BdTc=BmTc And BdYy=BmYy And BdChr=BmChr And BdNo=BmNo " + ws_Join + wCnd
  
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt  '****** Sachin 3.01 ctMaxDopOpt added
  
  '*** Jay 2.13(CT) ***
  'pg.19 - Bag typ & Bag chr not applicable in Stk Reco
  If UCase(adc.MenuCd) <> UCase("RepJTStkReco") Then
    'Uma *** blank <> 'D'
    If adc("UoBagTyp") = "F" Then
      Rep.TxtBagTypAndChr.SetText "(Flute Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
    ElseIf adc("UoBagTyp") = "P" Then
      Rep.TxtBagTypAndChr.SetText "(Primary Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
    ElseIf adc("UoBagTyp") = "C" Then
      Rep.TxtBagTypAndChr.SetText "(Component Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
    ElseIf adc("UoBagTyp") = "D" Then
      Rep.TxtBagTypAndChr.SetText "(PD Bags" & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
    End If
  End If
  
  'pg.19 - In Report, Heading & Bag No. label changed. Location & Gross wt not applicable in Stk Reco
  If UCase(adc.MenuCd) = UCase("RepJTStkReco") Then
    Rep.TxtHead.SetText "Stock Reconciliation"
    Rep.hBagNo.SetText "StkNum"
    Rep.hOthLoc.Suppress = True
    Rep.BdOthLoc1.Suppress = True
    Rep.hBrGrWt.Suppress = True
    Rep.BrGrWt.Suppress = True
    Rep.hOmCmCd.Suppress = True
    Rep.OmCmCd.Suppress = True
  End If
  
  'pg.19 - Company code added.
  Rep.TxtVchNoFr.SetText adc("UoCoCdFr") & "/" & adc("UoTTcFr") & "/" & adc("UoTYyFr") & "/" & adc("UoTChrFr") & "/" & adc("UoBrNoFr")
  Rep.TxtVchNoTo.SetText adc("UoCoCdFr") & "/" & adc("UoTTcTo") & "/" & adc("UoTYyTo") & "/" & adc("UoTChrTo") & "/" & adc("UoBrNoTo")
  If adc("UoTTcFr") = "" And adc("UoTYyFr") = "" And adc("UoTChrFr") = "" And adc("UoBrNoFr") = 0 Then Rep.TxtVchNoFr.Suppress = True
  If adc("UoTTcTo") = "" And adc("UoTYyTo") = "" And adc("UoTChrTo") = "" And adc("UoBrNoTo") = 0 Then Rep.TxtVchNoTo.Suppress = True
    
'  '***(Jen 2.11)
'  If adc("UoBagTyp") <> "" Then Rep.TxtBagTypAndChr.SetText IIF(adc("UoBagTyp") = "Y", "(Flute Bags", "(Primary Bags") & IIF(adc("UoBChrFr") <> "", ",  Chr = " & adc("UoBChrFr"), "") & ")"
'  '***(Jen 2.11)
'*** Jay 2.13(CT) ***
  CRV_REP.DisplayGroupTree = False
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
  '*** Set the first Non Key Control as UoCoCdFr
  Set moCn = adc.Connection  '*** (Bef speed) bef 24/11/06
  Set adc.FirNKeyCtl = adc("UoPrtKeyFr")  'Sachin 3.02 - Speed Optimization
  'Set adc.FirNKeyCtl = adc("UoTTcFr")
   
   'pg.19 - Bag type and Bag Character filters are not applicable in Stock Reco
  If UCase(adc.MenuCd) = UCase("RepJTStkReco") Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChrFr").Visible = False
  End If
   
  Call InitProp(Me)
  
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  
  'pg.19 - For Stk Reco, Tc typ is "JRC" and Bag Reco, Tc typ is "RC"
  gs_TxnTcTyp = IIF(UCase(adc.MenuCd) = UCase("RepJTStkReco"), "JRC", "RC")
  '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
  Call SetGroupSort
  
End Sub

Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub

Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  '*** (Jen 2.11)
  Select Case UCase(IdName)
    Case Is = UCase("UoBChrFr")
      If adc.Mode = xNorm Then
        If adc("UoBagTyp") = "" Then Cancel = True: ErrMsg = "First Specify Flute Bag Y/N.": Exit Sub
        HlpList.PMCd "BCHR", "'" + adc("UoBagTyp") + "'": Exit Sub
      End If
  End Select
  '*** (Jen 2.11)
  
  
  SetProp Me, IdName, When
End Sub
Private Sub atxt_ReCalculate(Index As Integer, ByVal IdName As String, ByVal OldVal As Variant, ByVal RowNum As Integer, ByVal SourceName As String, ByVal SourceColName As String)
' **** Zubin 211 **** '
  Select Case UCase(IdName)
    Case Is = UCase("UOBCHRFR")
      'If ADC("UoBagTyp") = "N" Then
      adc("UOBCHRFR") = ""
  End Select
' **** Zubin 211 **** '
End Sub

' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added
Private Sub OptReco_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
End Select
End Sub
' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added

' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added
Private Sub OptReco_GotFocus(Index As Integer)
  DispMsg FraOptReco.ToolTipText, etInfo
End Sub
' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added

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
  adc("UOASOFDT") = moCn.SrvrDate
  If adc("UOYN") = "" Then adc("UOYN") = "Y"
   
  '*** Jay 2.13(CT) ***
  Dim wHeadChk As MwfLib.MDORowSet
  Set wHeadChk = moCn.OpenRes("Select hFluteBagYN, hCompBagsYn from Head where HCoCd= '" + gs_CoCd + "' and " + _
                                                    "HCd= '" + ctSelfCmCd + "' ")
  If wHeadChk!hFluteBagYN = "N" And wHeadChk!hCompBagsYn = "N" Then
    LblBagTyp.Visible = False
    adc("UoBagTyp").Visible = False
    LblUoBChrFr.Visible = False
    adc("UoBChrFr").Visible = False
    adc("UoBagTyp") = "P": adc("UoBChrFr") = ""
  End If
  Set wHeadChk = Nothing
  '*** Jay 2.13(CT) ***
  
'*** Jay 2.13(CT) ***
'  '*** (Jen 2.11)
'  If moCn.GetFldVal("Select hFluteBagYN from Head where HCoCd= '" + adc("UoCoCdFr") + "' and " + _
'                                                    "HCd= '" + ctSelfCmCd + "' ") = "N" Then
'    LblPrimFluteBags.Visible = False
'    adc("UoBagTyp").Visible = False
'    LblUoBChrFr.Visible = False
'    adc("UoBChrFr").Visible = False
'  End If
'  adc("UoBagTyp") = "N": adc("UoBChrFr") = ""
'  '*** (Jen 2.11)
'*** Jay 2.13(CT) ***

' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added
If adc("UoYN") = "Y" Then OptReco(0).Value = True
If adc("UoYN") = "N" Then OptReco(1).Value = True
' ******** Manali 3.5.0 - 24/11/08 - Radio Buttons Added
End Sub
Private Sub DispCoNm()
  Dim ws_HName As String
  ws_HName = moCn.GetFldVal("Select HName from Head where HCoCd = '" + adc("UoCoCdFr") + "' and " + _
                                                      "HCd = '" + ctSelfCmCd + "' ")
  Rep.wCoCd.SetText ws_HName
  Rep.wDt.SetText "As On Date : " + Format(adc("UOASOFDT"), DtShortStr())   '6.1
  
  'pg.19 - All/Differential Bags/Stocks
  If UCase(adc.MenuCd) <> UCase("RepJTStkReco") Then
    Rep.wAllDiff.SetText IIF(adc("UOYN") = "Y", "All Bags", "Differential Bags")
  Else
    Rep.wAllDiff.SetText IIF(adc("UOYN") = "Y", "All Stocks", "Differential Stocks")
  End If
    
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText ws_HName
  If moCn.RecSeek("Select 'x' From Head Where HCoCd='" + gs_CoCd + "' and HCd='" + ctSelfCmCd + "' and HShowRepLogo='Y' ") Then
    Rep.RHMain.UnderlaySection = True: Rep.wCoCd.Suppress = True
  Else
    Rep.RHLogo.Suppress = True
  End If
  ' ***** Manali 3.8.0 - Comp Logo in Report
  
' ******** Mnaali - Before 3.2.2
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
