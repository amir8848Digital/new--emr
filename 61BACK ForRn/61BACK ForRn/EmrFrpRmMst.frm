VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.1#0"; "mwfctlforspeed9.ocx"
Begin VB.Form EmrFrpRmMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raw Material/ Custom Master"
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
      Left            =   10860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9630
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
      Left            =   4965
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
      Left            =   6180
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
      Width           =   15285
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
         TabPicture(0)   =   "EmrFrpRmMst.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_tabrep(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Group Sort"
         TabPicture(1)   =   "EmrFrpRmMst.frx":001C
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
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   635
            End
         End
         Begin VB.Frame fra_tabrep 
            BorderStyle     =   0  'None
            Height          =   9070
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   15105
            Begin VB.Frame FraOptBag 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1770
               TabIndex        =   34
               ToolTipText     =   "Show Open Or Closed Or All Bags"
               Top             =   1750
               Width           =   3465
               Begin VB.OptionButton OptBag 
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
                  TabIndex        =   10
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.OptionButton OptBag 
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
                  Left            =   1380
                  TabIndex        =   11
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
                  TabIndex        =   12
                  ToolTipText     =   "Show Open Or Closed Or All Bags"
                  Top             =   0
                  Width           =   1005
               End
            End
            Begin VB.Frame FraOptRmCd 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1740
               TabIndex        =   33
               ToolTipText     =   "Select Whether To Show Custom Code Or Rm Code"
               Top             =   570
               Width           =   5595
               Begin VB.OptionButton OptRmCd 
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
                  Index           =   1
                  Left            =   2460
                  TabIndex        =   2
                  ToolTipText     =   "Select Whether To Show Custom Code Or Rm Code"
                  Top             =   0
                  Width           =   1815
               End
               Begin VB.OptionButton OptRmCd 
                  Caption         =   "Customs Code"
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
                  TabIndex        =   1
                  ToolTipText     =   "Select Whether To Show Custom Code Or Rm Code"
                  Top             =   0
                  Width           =   1815
               End
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   8
               Left            =   4200
               TabIndex        =   8
               ToolTipText     =   "Enter Raw Material Code To Range"
               Top             =   1425
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDTO"
               IdName          =   "UORMCDTO"
               CmpStr          =   "RmCd <="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   15
               Left            =   6660
               TabIndex        =   9
               ToolTipText     =   "Enter Raw Material Code Selection"
               Top             =   1425
               Width           =   5640
               _ExtentX        =   9948
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMCDSEL"
               IdName          =   "UORMCDSEL"
               CmpStr          =   "RmCd In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   1
               Left            =   4200
               TabIndex        =   5
               ToolTipText     =   "Enter Raw Material Sub Category To Range"
               Top             =   1140
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
               Index           =   2
               Left            =   6660
               TabIndex        =   6
               ToolTipText     =   "Enter Raw Material Sub Category Selection"
               Top             =   1140
               Width           =   5640
               _ExtentX        =   9948
               _ExtentY        =   503
               MaxLength       =   255
               DataField       =   "UORMSCTGSEL"
               IdName          =   "UORMSCTGSEL"
               CmpStr          =   "RmSCtg In"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   3
               Left            =   1740
               TabIndex        =   3
               ToolTipText     =   "Enter Raw Material Category"
               Top             =   855
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   503
               MaxLength       =   2
               DataField       =   "UORMCTGFR"
               IdName          =   "UORMCTGFR"
               CmpStr          =   "RmCtg ="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   0
               Left            =   1740
               TabIndex        =   4
               ToolTipText     =   "Enter Raw Material Sub Category From Range"
               Top             =   1140
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
               Left            =   1740
               TabIndex        =   7
               ToolTipText     =   "Enter Raw Material Code From Range"
               Top             =   1425
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   503
               MaxLength       =   16
               DataField       =   "UORMCDFR"
               IdName          =   "UORMCDFR"
               Data            =   "ABCDEWABCDEW"
               CmpStr          =   "RmCd >="
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   4
               Left            =   1740
               TabIndex        =   0
               ToolTipText     =   "Specify Whether To Show Custom Code Or Rm Code (Enter 'Y' For Custm Cd/ 'N' For Rm Cd)"
               Top             =   570
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN"
               IdName          =   "UOYN"
            End
            Begin MwfCtl.MWCTL_MED atxt 
               Height          =   285
               Index           =   5
               Left            =   5370
               TabIndex        =   35
               ToolTipText     =   "Show Open Bags (Yes / No / All)"
               Top             =   1750
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               Enabled         =   0   'False
               DataType        =   4
               DataField       =   "UOYN1"
               IdName          =   "UOYN1"
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Valid YN"
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
               TabIndex        =   36
               Top             =   1750
               Width           =   1425
            End
            Begin VB.Label LblPic 
               BackStyle       =   0  'Transparent
               Caption         =   "Show Custm Cd"
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
               TabIndex        =   32
               ToolTipText     =   "Location"
               Top             =   570
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Category"
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
               Left            =   75
               TabIndex        =   31
               ToolTipText     =   "Location"
               Top             =   855
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Category"
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
               TabIndex        =   30
               Top             =   1140
               Width           =   1605
            End
            Begin VB.Label ALBL 
               BackStyle       =   0  'Transparent
               Caption         =   "Code"
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
               Left            =   75
               TabIndex        =   29
               Top             =   1425
               Width           =   1605
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
               TabIndex        =   28
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
               Left            =   4200
               TabIndex        =   27
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
               TabIndex        =   26
               Top             =   120
               Width           =   1335
            End
         End
      End
   End
End
Attribute VB_Name = "EmrFrpRmMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rep As New EmrRepRmMst
Dim moCn As MwfLib.MDOConnection   '*** (Bef speed) bef 24/11/06
Private Sub SetGroupSort()
  GRP_REP.Add "Sub Ctg", "RmSCtg", "", "RmSCtg", "hRmSCtg", "", "", "", "(Select PDesc From Param Where PTyp= 'RMSCTG' and PMCd = RmCtg and PSCd= RmSCtg) "
End Sub

Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
If adc("UoYN") = "Y" Then OptRmCd(0).Value = True
If adc("UoYN") = "N" Then OptRmCd(1).Value = True
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

'4.1.0.0
If adc("UoYN1") = "Y" Then
  OptBag(0).Value = True
ElseIf adc("UoYN1") = "N" Then
  OptBag(1).Value = True
ElseIf adc("UoYN1") = "" Then
  OptBag(2).Value = True
End If

End Sub

Private Sub adc_setreprecsource()
  '*** Report Sql ***
  '  Select space(30) as Grp1, space(30) as Grp2, space(30) as Grp3,
  '  space(30) as DGrp1, space(30) as DGrp2, space(30) as DGrp3,
  '  space(30) as G1Desc, space(30) as G2Desc, space(30) as G3Desc,
  '  RmCtg, RmCd, RmSCtg, RmDesc, RmUom, RmCol, RmKt, RmZCd,
  '  RmSpGrav, RmPurityRt, RmPurityWt, RmIntQly, RmCmCd,
  '  RmQw , RmQwCst, RmBaseCd, RmBasePurityWt, RmValidYn,
  '  RaCmpCtg , RaCmpSCtg, RaCmpCd, RaPerGmWt, RmFndWt
  '  from RmMst, RmAlyCmp where 1= 2
  '*** Report Sql ***
  
  Dim grpflds As String, wrepcnd As String, wCnd As String, wSqlStrg As String, wStr As String
  
  Set Rep = Nothing
  Set adc.RepSource = Rep
  Call DispCoNm
  Call SetMwName(Rep)   ' ***** Manali 3.6.1 - Microway footer
  Call GRP_REP.Gen3LRep
  grpflds = GRP_REP.GrpFldLst
    
  'If adc.MenuCd = UCase("RepRm") Then
  If adc("UoYN") = "N" Then
    wStr = " RmZ= 'N' "
    Rep.TxtHead.SetText "Raw Material Master Listing"
    Rep.TxtRmZCd.Suppress = False
    Rep.FldRmZCd.Suppress = False
  'ElseIf adc.MenuCd = UCase("RepRmZ") Then
  ElseIf adc("UoYN") = "Y" Then
    wStr = " RmZ= 'Y' "
    Rep.TxtHead.SetText "Custom Code Listing"
    Rep.TxtRmZCd.Suppress = True
    Rep.FldRmZCd.Suppress = True
  End If
  
  If adc("UoRmCtgFr") = "X" Then
    'Rep.hRmQw.Suppress = False
    'Rep.FldRmQw.Suppress = False
    Rep.hRmBaseCd.Suppress = False
    Rep.FldRmBaseCd.Suppress = False
    
    '*** (Jen 3.01)
    Rep.hRmSpGrav.Suppress = True: Rep.FldRmSpGrav.Suppress = True
    Rep.hRmPurityRt.Suppress = True: Rep.FldRmPurityRt.Suppress = True
    Rep.hRmPuityWt.Suppress = True: Rep.FldRmPuityWt.Suppress = True
    Rep.hRmFndWt.Suppress = False: Rep.FldRmFndWt.Suppress = False
    '*** (Jen 3.01)
    ' ***** Manali 3.6.0 - 29/08/09 - RmSTWGrp added
    Rep.hRmSTWGrp.Suppress = True: Rep.FldRmSTWGrp.Suppress = True
    Rep.FldRmAlyCd.Suppress = True
  Else
    'Rep.hRmQw.Suppress = True
    'Rep.FldRmQw.Suppress = True
    Rep.hRmBaseCd.Suppress = True
    Rep.FldRmBaseCd.Suppress = True
    
    '*** (Jen 3.01)
    Rep.hRmSpGrav.Suppress = False: Rep.FldRmSpGrav.Suppress = False
    Rep.hRmPurityRt.Suppress = False: Rep.FldRmPurityRt.Suppress = False
    Rep.hRmPuityWt.Suppress = False: Rep.FldRmPuityWt.Suppress = False
    'Rep.hRmFndWt.Suppress = True: Rep.FldRmFndWt.Suppress = True
    '*** (Jen 3.01)
    ' ***** Manali 3.6.0 - 29/08/09 - RmSTWGrp added
    ' **** Manali 3.8.0 - RmAlyCd added
    If adc("UoRmCtgFr") = "D" Or adc("UoRmCtgFr") = "C" Then
      Rep.hRmSTWGrp.SetText "STW Group"
    Else
      Rep.hRmSTWGrp.SetText "Aly Cd"
    End If
    Rep.FldRmSTWGrp.Suppress = False:  Rep.FldRmAlyCd.Suppress = False
  End If
  
  wrepcnd = adc.RepCond
  wCnd = IIF(wrepcnd <> "", " and " + wrepcnd, "")

  '4.1.0.0
  If adc("UOYN1") = "Y" Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " RmValidYN='Y' "
  ElseIf adc("UOYN1") = "N" Then
    wCnd = wCnd + IIF(wCnd <> "", " and ", "") + " RmValidYN='N' "
  End If

  '****** (Jen 3.01) (added RmFndWt)
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date:   20/01/2004
' ***** Manali 3.6.0 - 29/08/09 = RmSTWGrp added
' ***** Manali 3.8.0 - RmAlyCd added
  wSqlStrg = "Select " + grpflds + ", " + _
             "RmCtg, RmCd, RmSCtg, RmDesc, RmUom, RmCol, RmKt, RmZCd, RmSpGrav, " + _
             "RmPurityRt, RmPurityWt, RmIntQly, RmCmCd, RmQw, RmQwCst, RmBaseCd, RmBasePurityWt, RmValidYn, RmSTWGrp, RmAlyCd, " + _
             "IsNull(RaCmpCtg, '') as RaCmpCtg, IsNull(RaCmpSCtg, '') as RaCmpSCtg, " + _
             "IsNull(RaCmpCd, '') as RaCmpCd, IsNull(RaPerGmWt, 0) as RaPerGmWt, RmFndWt " + _
             "From RmMst Left Outer Join RmAlyCmp on RaCd= RmCd where " + IIF(gs_Partition = ctCurrPrtn, " RmPrtKey='" + ctCurrPrtn + "' and ", "") + wStr + wCnd
' ###########################################  Manoj  ###########################################
  adc.RepRecSource = wSqlStrg + ctMaxDopOpt '****** Sachin 3.01 added ctMaxDopOpt
  '578.19 calling function to change Rmcode font size
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmCd}", 8)
  GRP_REP.SetFormula Rep, "wRmCdFontSz", GetRmcdFontSz("{rdo.RmZCd}", 8)
  
  CRV_REP.DisplayGroupTree = False
  
  Call DispReport(adc) '****** Sachin 3.02 - Display Report ******
End Sub
Private Sub OptBag_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN1") = "Y"
Case Is = 1
  adc("UoYN1") = "N"
Case Is = 2
  adc("UoYN1") = ""
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
    Set moCn = adc.Connection    '*** (Bef speed) bef 24/11/06
    'Set adc.FirNKeyCtl = adc("UoYn")
    Set adc.FirNKeyCtl = adc("UoRmCtgFr")
    Call InitProp(Me)
    
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    gb_RmCtgFor = True
    'If UCase(adc.MenuCd) = UCase("RepRm") Then gb_RmZ = False
    'If UCase(adc.MenuCd) = UCase("RepRmZ") Then gb_RmZ = True
    '*** Parameters Used For Setting Properties Of UsrOpt Fields ***
    
    Call SetGroupSort
End Sub
Private Sub ATXT_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  Call SetProp(Me, IdName, Valid, Cancel, ErrMsg)
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
  If adc("UoYN") = "N" Then gb_RmZ = False
  If adc("UoYN") = "Y" Then gb_RmZ = True
  
  SetProp Me, IdName, When
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptRmCd_Click(Index As Integer)
Select Case Index
Case Is = 0
  adc("UoYN") = "Y"
Case Is = 1
  adc("UoYN") = "N"
End Select
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added
Private Sub OptRmCd_GotFocus(Index As Integer)
  DispMsg FraOptRmCd.ToolTipText, etInfo
End Sub
' ****** Manali 3.5.0 - 03/11/08 - Radio Buttons Added

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
  
  ' ***** Manali 3.8.0 - Comp Logo in Report
  Rep.wCoCdLogo.SetText gs_CoNm
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
