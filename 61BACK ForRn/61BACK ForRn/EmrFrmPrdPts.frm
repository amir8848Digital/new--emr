VERSION 5.00
Object = "{AB2CC32A-E652-42A2-B341-26ECE02F6386}#2.0#0"; "MwfCtl.ocx"
Begin VB.Form EmrFrmPrdPts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Points Master"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MwfCtl.MWCTL_BTN1 CmdCP 
      Height          =   405
      Left            =   5310
      TabIndex        =   11
      ToolTipText     =   "Option To Copy The Rates From Another Customer"
      Top             =   9390
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      BackColor       =   -2147483633
      MaskColor       =   8454143
      Caption         =   "&Copy"
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
      Height          =   405
      Left            =   6990
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   714
   End
   Begin MwfCtl.MWCTL_ADC ADC 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   0
      TabIndex        =   19
      Top             =   9390
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   979
      AllowAdd        =   0   'False
      AllowDelete     =   0   'False
      BtnWidth        =   885
      BtnHeight       =   435
      BtnMaskColor    =   8454143
      FontSz          =   8.25
   End
   Begin VB.Frame FRA_KEYS 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   -36
      TabIndex        =   21
      Top             =   24
      Width           =   11700
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   2
         ToolTipText     =   "Enter From Design Production Category"
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WPPPRDCTGFR"
         CmpStr          =   "PpPrdCtg >= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   3
         Left            =   3255
         TabIndex        =   3
         ToolTipText     =   "Enter To Design Production Category"
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WPPPRDCTGTO"
         CmpStr          =   "PpPrdCtg <= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   4
         Left            =   5205
         TabIndex        =   4
         ToolTipText     =   "Enter From Setting Type"
         Top             =   0
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPPSETTYPFR"
         CmpStr          =   "PpSetTyp >= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   2
         Left            =   5205
         TabIndex        =   5
         ToolTipText     =   "Enter To Setting Type"
         Top             =   285
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPPSETTYPTO"
         CmpStr          =   "PpSetTyp <= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   5
         Left            =   855
         TabIndex        =   1
         ToolTipText     =   "Enter Sort Option"
         Top             =   285
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPPSRT"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   7
         Left            =   855
         TabIndex        =   0
         ToolTipText     =   "Enter Production Type"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         MaxLength       =   5
         DataType        =   4
         IdName          =   "WPPTYP"
         CmpStr          =   "PpTyp= "
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   8
         Left            =   10095
         TabIndex        =   8
         ToolTipText     =   "Enter From Process"
         Top             =   0
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPPLOCFR"
         CmpStr          =   "PpLoc >="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   9
         Left            =   10095
         TabIndex        =   9
         ToolTipText     =   "Enter To Process"
         Top             =   285
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         MaxLength       =   8
         DataType        =   4
         IdName          =   "WPPLOCTO"
         CmpStr          =   "PpLoc <="
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   10
         Left            =   7230
         TabIndex        =   6
         ToolTipText     =   "Enter From Design"
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WPPDMCDFR"
      End
      Begin MwfCtl.MWCTL_MED ATXT 
         Height          =   285
         Index           =   17
         Left            =   7230
         TabIndex        =   7
         ToolTipText     =   "Enter To Design"
         Top             =   285
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   15
         DataType        =   4
         IdName          =   "WPPDMCDTO"
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "       To"
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
         Index           =   17
         Left            =   6450
         TabIndex        =   35
         Top             =   285
         Width           =   735
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dsg Fr"
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
         Index           =   16
         Left            =   6450
         TabIndex        =   34
         Top             =   0
         Width           =   735
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prc Fr"
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
         Index           =   9
         Left            =   9330
         TabIndex        =   30
         Top             =   0
         Width           =   705
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "      To"
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
         Index           =   8
         Left            =   9330
         TabIndex        =   29
         Top             =   285
         Width           =   705
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "PrdTyp"
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
         Index           =   7
         Left            =   60
         TabIndex        =   28
         Top             =   0
         Width           =   795
      End
      Begin VB.Label ALBL 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort On"
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
         Index           =   5
         Left            =   60
         TabIndex        =   27
         Top             =   285
         Width           =   825
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "           To"
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
         Index           =   4
         Left            =   4140
         TabIndex        =   26
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SetTyp Fr"
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
         Index           =   3
         Left            =   4140
         TabIndex        =   25
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "            To"
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
         Left            =   2130
         TabIndex        =   24
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label ALBL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PrdCtg Fr"
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
         Left            =   2130
         TabIndex        =   23
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.Frame FRA_NKEYS 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9360
      Left            =   -120
      TabIndex        =   20
      Top             =   0
      Width           =   12465
      Begin VB.Frame FraCp 
         Height          =   1035
         Left            =   2715
         TabIndex        =   12
         Top             =   8325
         Visible         =   0   'False
         Width           =   6015
         Begin MwfCtl.MWCTL_BTN1 CmdCpGo 
            Height          =   285
            Left            =   5325
            TabIndex        =   17
            ToolTipText     =   "Press This Button To Copy"
            Top             =   675
            Width           =   585
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483633
            MaskColor       =   8454143
            Caption         =   "&Go"
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
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   14
            Left            =   1155
            TabIndex        =   13
            ToolTipText     =   "Enter Copy Source Location"
            Top             =   390
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPPPLOCFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   16
            Left            =   1155
            TabIndex        =   14
            ToolTipText     =   "Enter Target Location"
            Top             =   675
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            MaxLength       =   8
            DataType        =   4
            IdName          =   "WCPPPLOCTO"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   1
            Left            =   4020
            TabIndex        =   15
            ToolTipText     =   "Enter Source Sub Process"
            Top             =   360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCPPPSUBPRCFR"
         End
         Begin MwfCtl.MWCTL_MED ATXT 
            Height          =   285
            Index           =   6
            Left            =   4020
            TabIndex        =   16
            ToolTipText     =   "Enter Target Sub Process"
            Top             =   645
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            MaxLength       =   15
            DataType        =   4
            IdName          =   "WCPPPSUBPRCTO"
         End
         Begin VB.Label LblMain 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Copy From"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   -30
            TabIndex        =   37
            Top             =   60
            Width           =   6015
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Sub Prc"
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
            Left            =   2475
            TabIndex        =   36
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Sub Prc"
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
            Index           =   6
            Left            =   2475
            TabIndex        =   33
            Top             =   645
            Width           =   1635
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Prc"
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
            Index           =   15
            Left            =   75
            TabIndex        =   32
            Top             =   375
            Width           =   1245
         End
         Begin VB.Label ALBL 
            BackStyle       =   0  'Transparent
            Caption         =   "Target Prc"
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
            Index           =   13
            Left            =   75
            TabIndex        =   31
            Top             =   660
            Width           =   1185
         End
      End
      Begin VB.Frame FraNKeyAll 
         BorderStyle     =   0  'None
         Height          =   9225
         Left            =   150
         TabIndex        =   22
         Top             =   90
         Width           =   12270
         Begin MwfCtl.MWCTL_FLX GrdPP 
            Height          =   8670
            Left            =   0
            TabIndex        =   10
            Top             =   570
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   15293
            Cols            =   14
            colname1        =   "PPLOC"
            heading1        =   "Process        "
            datafld1        =   "PpLoc"
            datatype1       =   4
            recalcparent1   =   "GrdPP(PPSUBPRC)"
            maxlength1      =   8
            tooltiptext1    =   "Enter Process"
            colname2        =   "PPSUBPRC"
            heading2        =   "Sub Prc"
            datafld2        =   "PpSubPrc"
            datatype2       =   4
            recalcon2       =   "GrdPP(PPLOC)"
            maxlength2      =   8
            tooltiptext2    =   "Enter Sub Process"
            colname3        =   "PPPRDCTG"
            heading3        =   "Prd Ctg "
            datafld3        =   "PpPrdCtg"
            datatype3       =   4
            maxlength3      =   5
            tooltiptext3    =   "Enter Design Production Category"
            colname4        =   "PPSETTYP"
            heading4        =   "Set Typ        "
            datafld4        =   "PpSetTyp"
            datatype4       =   4
            maxlength4      =   8
            tooltiptext4    =   "Enter Setting Type"
            colname5        =   "WNOTUSED"
            datatype5       =   4
            maxlength5      =   3
            style5          =   0
            colname6        =   "PPDMCD"
            heading6        =   "Design Cd                       "
            datafld6        =   "PpDmCd"
            datatype6       =   4
            maxlength6      =   15
            tooltiptext6    =   "Enter Design Code"
            colname7        =   "PPRMCTG"
            heading7        =   "Rm Ctg"
            datafld7        =   "PpRmCtg"
            maxlength7      =   2
            style7          =   0
            tooltiptext7    =   "Enter Raw Material Category"
            colname8        =   "PPRMSCTG"
            heading8        =   "Rm SCtg"
            datafld8        =   "PpRmSCtg"
            maxlength8      =   5
            style8          =   0
            tooltiptext8    =   "Enter Raw Material Sub Category"
            colname9        =   "PPFRRMWT"
            heading9        =   "From Rm Wt"
            datafld9        =   "PpFrRmWt"
            datatype9       =   2
            mask9           =   "####0.000"
            maxlength9      =   9
            tooltiptext9    =   "Enter From Rm Weight"
            colname10       =   "PPTORMWT"
            heading10       =   "To Rm Wt     "
            datafld10       =   "PpToRmWt"
            datatype10      =   2
            mask10          =   "####0.000"
            maxlength10     =   9
            tooltiptext10   =   "Enter From Rm Weight"
            colname11       =   "PPPTS"
            heading11       =   "Points    "
            datafld11       =   "PpPts"
            datatype11      =   2
            mask11          =   "###0.00"
            maxlength11     =   7
            tooltiptext11   =   "Enter Production Points"
            colname12       =   "PPPRTKEY"
            heading12       =   "Prtn "
            datafld12       =   "PpPrtKey"
            maxlength12     =   1
            style12         =   0
            colname13       =   "PPSKL"
            heading13       =   "Skill"
            datafld13       =   "PpSkl"
            datatype13      =   4
            maxlength13     =   2
            style13         =   0
            tooltiptext13   =   "Enter Skill"
         End
      End
   End
End
Attribute VB_Name = "EmrFrmPrdPts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- General comments on the Entry  ----
' X-'', A-PrdPts, GrdPp
' The grid properties are set by the TcStructure Table.
'
' FraCp, CmdCp, CmdCpGo, CpyPoints, wCpPpPrdCtgFr, wCpPpPrdCtgTo, wCpPpSetTypFr,...wCpPpLocTo
'          all part of Copy routine.
'
' private Sub GrdPropSet, called from AdcRecsource, sets all properties for each grid col.
'     reads the TcStru table, assumes records for all fields will be avlbl for Tc/Ctg Level
'
' private Sub CmpString is used to set the values for Compare String property,
'            passing 'K'ey Fields, 'C'opy
'
' private fn ColSty, DataTyp, to set the Col Style, & Datatype property.
'
'-----
Option Explicit
Dim moCn As MwfLib.MDOConnection '*** (Bef speed) bef 24/11/06
Dim mPpTyp As String, mPpSrt As String
Dim ms_MenuCd As String '3.11.0
Dim ms_DmTcTyp As String
'*** For the Toggle Options of Command Buttons
Enum en_PpFra
  Cpy = 0
End Enum

Private Sub ADC_Load()
  '*** Set the First Key Control property to the working field wPpTyp
  '*** Set the First Non Key Control property to the grid GrdPp
  '*** Set the Child Property Of ADC to GrdPp
  '*** Set the Previous and Next Control Properties for the Grid GrdPp as ADC
  '*** Set the Hot Key for the Grid GrdPP as 'A'
  '*** initialise the default values for the working fields wPpTyp and wPpSrt

  Set moCn = ADC.Connection  '*** (Bef speed) bef 24/11/06
  ms_MenuCd = UCase(ADC.MenuCd)

  'New Menu 'BL' introduced in 3.11.0 (Hard Coded Prodn Type = 'BL')
  If ms_MenuCd = "PRCLMST" Then
    mPpTyp = "BL"
    Me.Caption = "Process Loss Master"
    Set ADC.FirKeyCtl = ADC("wPpPrdCtgFr")
    ALBL(1).Caption = "DmCtg Fr": ALBL(3).Visible = False: ALBL(4).Visible = False
    ALBL(5).Visible = False: ADC("wPpSrt").Visible = False
    ADC("wPpSetTypFr").Visible = False: ADC("wPpSetTypTo").Visible = False
    ms_DmTcTyp = "DM"
  Else
  ' there are only 3 menu code used for this form - still just to keep it clean checking again below menucd but there is no other possibility
   If ms_MenuCd = "PRDPTS" Or ms_MenuCd = "PRDPTSSKETCH" Then
      mPpTyp = ""
      Set ADC.FirKeyCtl = ADC("wPpTyp")
      ADC("wPpTyp").Enabled = True
      ms_DmTcTyp = IIF(ms_MenuCd = "PRDPTS", "DM", "SM")
   End If
  End If

  Set ADC.FirNKeyCtl = GrdPP
  ADC.Child = "GrdPp"

  Set GrdPP.PrevCtl = ADC
  Set GrdPP.NextCtl = ADC
  GrdPP.HotKey = "A"

  mPpSrt = ""
  
  '***** Sachin 3.02.0 - Partition Key
  GrdPP.ColProp("PpPrtKey").Style = IIF(gs_Partition = ctCurrPrtn, fgcinvisible, fgcReadOnly)
End Sub
Private Sub ADC_KeyWhen()
  '*** pass 'K' as the parameter for the CmpString procedure. This sets the CmpStr value of
      'the fields in the copy frame as blank and sets the proper cmpstr values for fileds in
      'the key frame (i.e. the range selection fields)
  '*** set the default values for PpTyp and Sort Option from the previous entry
  '*** Disable All command buttons
  '*** Hide all frames except Frame FraNKeyAll

  Call CmpString("K")     ' to set the Compare String Property of the ADC controls when the Focus is in the Key Frame
 
  If ms_MenuCd = "PRCLMST" Then
    ADC("wPpTyp") = mPpTyp
    ADC("wPpTyp").Enabled = False
    Call GrdPropSet
  Else
    If mPpTyp <> "" Then ADC("wPpTyp") = mPpTyp
  End If
  
  If mPpSrt <> "" Then ADC("wPpSrt") = mPpSrt

  Call EnaDisaCmds(True)
  Call HideAllFras
End Sub
Private Sub ADC_KeyValid(Cancel As Boolean, ErrMsg As String)
  '*** Check whether the PpTyp and the Sort Option is valid
      '(Sort option can be blank in which case the default sorting is on PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts)
 
  '*** Checking Key values in Key Valid as they are copied in Init
  Call atxt_FldChange(0, UCase("wPpTyp"), Cancel, ErrMsg, ADC("wPpTyp"))
  
  Cancel = True
  Cancel = (ADC("wPpSrt") <> "") And _
           (Not moCn.RecSeek("Select PMCd From Param Where " + _
                             "PTyp = 'SRTPP' And PMCd = '" + ADC("wPpSrt") + "'"))
  If Cancel = True Then ErrMsg = "Invalid Sort Option": Exit Sub
End Sub
Private Sub ADC_NonKeyWhen(pr_Cancel As Boolean, pr_ErrMsg As String)
  '*** Store the PpTyp, Sort Option for the next entry
  '*** Enable all the command buttons
  
  mPpTyp = ADC("wPpTyp")
  mPpSrt = ADC("wPpSrt")
  
  Call EnaDisaCmds(False)
  
  
End Sub
Private Sub ATXT_FldWhen(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String)
'*** Set Helps For Each Field
  
  Select Case IdName
'*** Fields of The Key Frame & The Copy Frame
  Case Is = UCase("wPpTyp")
    Call HlpList.PMCd("PRDTYP")
  Case Is = UCase("wPpSrt")
    Call HlpList.PMCd("SRTPP")
  Case Is = UCase("wPpPrdCtgFr"), UCase("wPpPrdCtgTo") ', UCase("wCpPpPrdCtgFr"), UCase("wCpPpPrdCtgTo") '***** Changed By Shilpa Emr 206 29/10/03 ******
    If ms_MenuCd <> "PRCLMST" Then
        Call HlpList.PMCd("PRDCTG", , , , False)
    Else
        Call HlpList.PMCd("DMCTG")  '3.11.0
    End If
  Case Is = UCase("wPpSetTypFr"), UCase("wPpSetTypTo") ', UCase("wCpPpSetTypFr"), UCase("wCpPpSetTypTo") '***** Changed By Shilpa Emr 206 29/10/03 ******
    Call HlpList.PSCd("LABSCD", "SET", , False)
  Case Is = UCase("wPpDmCdFr"), UCase("wPpDmCdTo") ', UCase("wCpPpDmCdFr"), UCase("wCpPpDmCdTo") '***** Changed By Shilpa Emr 206 29/10/03 ******
    Call HlpList.DmCd(ms_DmTcTyp)
  Case Is = UCase("wPpLocFr"), UCase("wPpLocTo"), UCase("wCpPpLocFr")
  'uma ***** p replaced by R
    Call HlpList.Loc("", "'R'", , , False)
  '*** Jay 2.14 ***
  'uma ***** p replaced by R
  Case Is = UCase("wCpPpLocTo")
    Call HlpList.Loc("", "'R'")
  '*** Jay 2.14 ***
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  Case Is = UCase("wCpPpSubPrcFr")
    Call HlpList.PSCd("SUBPRC", ADC("wCpPpLocFr"), , False)
  Case Is = UCase("wCpPpSubPrcTo")
    Call HlpList.PSCd("SUBPRC", ADC("wCpPpLocTo"))
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  End Select
End Sub
Private Sub atxt_FldChange(Index As Integer, ByVal IdName As String, Cancel As Boolean, ErrMsg As String, ByVal pv_NewValue As Variant)
  '*** wPpTyp       valid PMCd('PRDTYP')
  '*** wPpSrt       Blank or valid PMCd('SRTPP')
  '*** wCpyCustFr   Blank or valid CmCd from CustMst(CmCtg= 'C')
  '*** wCpyCustTo   Blank or valid CmCd from CustMst(CmCtg= 'C')
  
  Select Case UCase(IdName)
  Case Is = UCase("wPpTyp")
    '3.11.0 '3.12.0 - Name Change
    Cancel = Not (ms_MenuCd = "PRCLMST" And pv_NewValue = "BL") And Not moCn.RecSeek("Select PMCd From Param where PTyp= 'PRDTYP' " + _
                         "and PMCd = '" + pv_NewValue + "'")
    ErrMsg = "Invalid Production Type"
    If Cancel = True Then Exit Sub
    If ms_MenuCd = UCase("PrdPtsSketch") Then
      Cancel = pv_NewValue <> "GN"
      ErrMsg = "Type has to be GN for Production Points for Sketch"
   End If
    Exit Sub
  Case Is = UCase("wPpSrt")
    If pv_NewValue <> "" Then
      Cancel = Not moCn.RecSeek("Select PMCd From Param where PTyp= 'SRTPP' " + _
                           "and PMCd = '" + pv_NewValue + "'")
      ErrMsg = "Invalid Sort Option": Exit Sub
    End If
    
  '***** Changed By Shilpa Emr 206 29/10/03 ******  '3.12.0 - Name Change
  'uma *** changed P to R
  Case Is = UCase("wCpPpLocFr"), UCase("wCpPpLocTo")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select LocCd From Loc where " + _
                                    "LocTyp= 'R' and LocCd = '" + pv_NewValue + "'" + IIF(ms_MenuCd <> "PRCLMST", " And LocPrdTyp = '" + ADC("wPPTyp") + "'", "")) '3.11.0
    If Cancel = True Then ErrMsg = "Invalid Process ": Exit Sub
    
    
  Case Is = UCase("wCpPpSubPrcFr")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select PSCd From Param where PTyp= 'SUBPRC' " + _
                                    "and PMCd = '" + ADC("wCpPpLocFr") + "' and " + _
                                    "PSCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid From Sub Process": Exit Sub
    
  Case Is = UCase("wCpPpSubPrcTo")
    Cancel = (pv_NewValue <> "") And Not moCn.RecSeek("Select PSCd From Param where PTyp= 'SUBPRC' " + _
                                    "and PMCd = '" + ADC("wCpPpLocTo") + "' and " + _
                                    "PSCd = '" + pv_NewValue + "' ")
    If Cancel = True Then ErrMsg = "Invalid To Sub Process": Exit Sub
    
  'Case Is = UCase("wCpPpPrdCtgFr")
   ' Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PmCd From Param where " + _
             "PTyp= 'PRDCTG' and PmCd = '" + pv_NewValue + "'")
    ' ErrMsg = "Invalid Source Production Category": Exit Sub
  'Case Is = UCase("wCpPpPrdCtgTo")
   ' Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PmCd From Param where " + _
             "PTyp= 'PRDCTG' and PmCd = '" + pv_NewValue + "'")
    ' ErrMsg = "Invalid Target Production Category": Exit Sub
  'Case Is = UCase("wCpPpSetTypFr")
   ' Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param where " + _
             "PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = '" + pv_NewValue + "'")
    ' ErrMsg = "Invalid Source Setting Type": Exit Sub
  'Case Is = UCase("wCpPpSetTypTo")
   ' Cancel = pv_NewValue <> "" And Not moCn.RecSeek("Select PSCd From Param where " + _
             "PTyp= 'LABSCD' and PMCd= 'SET' and PSCd = '" + pv_NewValue + "'")
    ' ErrMsg = "Invalid Target Setting Type": Exit Sub
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  End Select
End Sub

Private Sub GrdPP_ColWhen(ByVal RowNum As Integer, ByVal ColName As String, Cancel As Boolean, ErrMsg As String, ByVal Action As MwfCtl.en_fgAction)
  '*** Set Helps For Each Field
  With GrdPP
  
    '****** Sachin 3.02.0 - Partition Locking
    If GrdPP.Value(RowNum, "PpPrtKey") <> ctCurrPrtn Then Cancel = True: ErrMsg = "Cannot Edit Previous Partition Records"
  
    Select Case ColName
    Case Is = UCase("PpLoc")
    'uma **** P replaced by R
      Call HlpList.Loc("", "'R'")
    Case Is = UCase("PpSubPrc")
      Call HlpList.PSCd("SUBPRC", .Value(RowNum, "PpLoc"))
    Case Is = UCase("PpPrdCtg")
      Call HlpList.PMCd(IIF(ms_MenuCd = "PRCLMST", "DMCTG", "PRDCTG"))      '3.11.2 - Name Change
    Case Is = UCase("PpSetTyp")
      Call HlpList.PSCd("LABSCD", "SET")
    Case Is = UCase("PpDmCd")
      Call HlpList.DmCd(ms_DmTcTyp)
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 03/12/2003
' Change: For PpRmCtg & PpRmCtg Help is Added.
    Case Is = UCase("PpRmCtg")
      Call HlpList.PMCd("RMCTG")
    Case Is = UCase("PpRmSCtg")
      Call HlpList.PSCd("RMSCTG", .Value(RowNum, "PpRmCtg"))
' ###########################################  Manoj  ###########################################
    
    'uma **** added new skill
    Case Is = UCase("PpSkl")
      Call HlpList.PMCd("SKL")
    End Select
  End With
End Sub
Private Sub GrdPP_ReCalc(ByVal RowNum As Integer, ByVal ColName As String, ByVal OldVal As Variant, ByVal SourceName As String, ByVal SourceColName As String)
  '*** PpSubPrc on PpLoc        Normal
      'PpSubPrc = Blank
  
  Select Case ColName
  Case Is = UCase("PpSubPrc")
    If GrdPP.Mode = fgmnorm Then
      GrdPP.Value(RowNum, "PpSubPrc") = ""
    End If
  End Select
End Sub
Private Sub GrdPP_ColChange(ByVal RowNum As Integer, ByVal ColName As String, ByVal NewValue As Variant, Cancel As Boolean, ErrMsg As String)
  '*** PpLoc     Blank or valid Loc(LocTyp= 'P')
  '*** PpSubPrc  Blank or valid PSCd from PSCd('SUBPRC', PpLoc)
  '*** PpPrdCtg  Blank or valid PMCd('PRDCTG')
  '*** PpSetTyp  Blank or valid PSCd('LABSCD', 'SET')
  '*** PpDmCd    Blank or valid DmCd('DM')
  '*** PpPts     >0
  
  'Dim ws_BlnkYN As String
  With GrdPP
    'ws_BlnkYN = moCn.GetFldVal("Select TsBlank from TcStr where TsTbl='RmRt' and " + _
                            "TsTyp='" + ADC("wRrTcTyp") + "' and " + _
                            "TsSTyp='" + ADC("wRrCtg") + "' and " + _
                            "TsIdName='" + ColName + "' ")
    
    Select Case UCase(ColName)
    'uma **** P replaced by R
    Case Is = UCase("PpLoc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek(" Select LocCd From Loc where " + _
                                    " LocTyp= 'R' and LocCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Process": Exit Sub
      
      If ms_MenuCd <> "PRCLMST" Then      '3.11.0   '3.11.2 - Name Change
        Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select LocCd From Loc where " + _
                                      "LocTyp= 'R' and LocCd = '" + NewValue + _
                                      "' And LocPrdTyp = '" + ADC("wPPTyp") + "'"))
        If Cancel = True Then ErrMsg = "Process Not Define For Production Type '" + ADC("wPPTyp") + "'": Exit Sub
      End If
      
    Case Is = UCase("PpSubPrc")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'SUBPRC' " + _
                                    "and PMCd = '" + .Value(RowNum, "PpLoc") + "' and " + _
                                    "PSCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Sub Process": Exit Sub
      
    Case Is = UCase("PpPrdCtg")
      '3.11.0   '3.11.2
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PmCd From Param where PTyp='" + _
                                    IIF(ms_MenuCd = "PRCLMST", "DMCTG", "PRDCTG") + "' and PmCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid " + IIF(ms_MenuCd = "PRCLMST", "Design", "Production") + " Category.": Exit Sub
      
    Case Is = UCase("PpSetTyp")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where PTyp= 'LABSCD' " + _
                                    "and PMCd= 'SET' and PSCd = '" + NewValue + "'"))
      If Cancel = True Then ErrMsg = "Invalid Setting Type": Exit Sub
      
    Case Is = UCase("PpDmCd")
      '****** Sachin 3.02.0
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select DmCd From DsgMst Where (DmTcTyp ='" + ms_DmTcTyp + "') " + _
                           "and DmCd = '" + NewValue + "' and DmSz= '' and DmPrtKey='" + ctCurrPrtn + "'"))
      If Cancel = True Then ErrMsg = "Invalid Design/Sketch Code": Exit Sub
      
' ###################################  Manoj #### Ver: 2.0.6  ###################################
' Date: 03/12/2003
' Change: For PpRmCtg & PpRmCtg validation is Added.
    Case Is = UCase("PpRmCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'RMCTG' and PMCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Category": Exit Sub
      Cancel = (NewValue <> "") And (NewValue <> "C" And NewValue <> "D")
      If Cancel = True Then ErrMsg = "Only 'C' and 'D' Raw Material Category are Allowed": Exit Sub
    Case Is = UCase("PpRmSCtg")
      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PSCd From Param where " + _
                                    "PTyp= 'RMSCTG' And PMCD = '" + .Value(RowNum, "PpRmCtg") + _
                                    "' and PSCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Raw Material Sub Category": Exit Sub
' ###########################################  Manoj  ###########################################
    Case Is = UCase("PpPts")
      If NewValue <= 0 Then Cancel = True: ErrMsg = "Prd Points Should Be Greater Than Zero": Exit Sub

    'uma **** added new skill
    Case Is = UCase("PpSkl")
      Cancel = (Not moCn.RecSeek("Select PMCd From Param where " + _
                                    "PTyp= 'Skl' and PMCd = '" + NewValue + "' "))
      If Cancel = True Then ErrMsg = "Invalid Skill": Exit Sub
    
'    Case Is = UCase("RrLabMCd")
'      If ws_BlnkYN = "N" And NewValue = "" Then _
'         Cancel = True: ErrMsg = "Enter Labour Main Code": Exit Sub
'      Cancel = (NewValue <> "") And (Not moCn.RecSeek("Select PMCd From Param where PTyp= 'LABMCD' " + _
'                           "and PMCd = '" + NewValue + "'"))
'      ErrMsg = "Invalid Labour Main Code": Exit Sub
    End Select
  End With
End Sub
Private Sub GrdPP_InitKey(ByVal RowNum As Integer)
  '*** Initialise Key Fields
  With GrdPP
    .Store "PpTyp", ADC("wPpTyp")
    .Store "PpDmTcTyp", ms_DmTcTyp
  End With
End Sub
Private Sub ADC_SetRecSource()
  '*** Since this is an 'A' Entry, there is no RecSource associated with ADC. So set the
      'DataMode property of ADC control as xdatUnbound

  ADC.DataMode = xdatUnbound
  If ms_MenuCd <> "PRCLMST" Then Call GrdPropSet
End Sub

Private Sub GrdPP_RowWhen(ByVal RowNum As Integer)

  '****** Sachin 3.02.0 - Partition Key
  If GrdPP.IsNew(RowNum) = True Then GrdPP.Value(RowNum, "PpPrtKey") = ctCurrPrtn

End Sub

Private Sub GrdPP_SetRecSource()
  '*** Set the Record Source of the Grid GrdPP
      'Default sort option is PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts
  
  Dim ws_Cnd As String, ws_OrdBy As String, Errstrg As String
  ws_Cnd = ADC.RepCond
  ws_Cnd = ws_Cnd + IIF(ws_Cnd <> "", " and ", "") + "(PpDmTcTyp='' or PpDmTcTyp='" + ms_DmTcTyp + "')"
  ws_Cnd = IIF(ws_Cnd <> "", " where " + ws_Cnd, "")
  If ADC("wPpSrt") = "" Then
    ws_OrdBy = " Order By PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts "
  Else
    ws_OrdBy = "Order By " + moCn.GetFldVal("Select PDesc225 from Param " + _
             "where PTyp= 'SRTPP' and PMCd= '" + ADC("wPpSrt") + "'")
  End If
  
  ' ****** Sachin 2.12 - 02-01-2006 - For Out of Memory Error
  If moCn.GetFldVal("Select Count(*) from PrdPts " + ws_Cnd + IIF(gs_Partition = ctCurrPrtn, IIF(Trim(ws_Cnd) = "", " Where ", " And ") + " PpPrtkey='" + ctCurrPrtn + "' ", "")) > 10000 Then
      'DispMsg "More Than 10000 Records to Fetch, Please Reduce Scope !!", etError
      GrdPP.RecSource = " Select * from PrdPts Where 1=2"
      MsgBox "More Than 10000 Records to Fetch, Please Reduce Scope !!"
      'adc.SetFocus
      Exit Sub
  End If
  ' ****** Sachin 2.12 - 02-01-2006 - For Out of Memory Error
  '****** Sachin 3.02.0 - Partition Key
  GrdPP.RecSource = " Select * from PrdPts " + ws_Cnd + IIF(gs_Partition = ctCurrPrtn, IIF(Trim(ws_Cnd) = "", " Where ", " And ") + " PpPrtkey='" + ctCurrPrtn + "' ", "") + ws_OrdBy

End Sub
Private Sub CmdCP_Click()
  
  '3.11.0   '3.12.0 - Name Change
  If ms_MenuCd = "PRCLMST" Then
    ALBL(0).Visible = False: ALBL(6).Visible = False
    ADC("wCpPpSubPrcFr").Visible = False: ADC("wCpPpSubPrcTo").Visible = False
  End If
  
  '*** Display or hide the Copy frame depending on the toggle satate
  Call DispFra(Cpy)
End Sub
Private Sub CmdCpGo_Click()
  '*** Check If Copy Is Possible, Call the Copy Routine and Close the Copy Frame
  
  '*** Checking if copy is possible
  '***** Changed By Shilpa Emr 206 29/10/03 ******
'  Select Case UCase(ADC("wPpTyp"))
'  Case Is = "GN", "WP"
'    If ADC("wCpPpPrdCtgFr") = "" Or ADC("wCpPpPrdCtgTo") = "" Then _
'       DispMsg "Enter Proper Source/ Target Prd Ctg", etError: Exit Sub
'    If ADC("wCpPpSetTypFr") <> "" Or ADC("wCpPpSetTypTo") <> "" Then _
'       DispMsg "Source/ Target Setting Typ Should Be Blank For This Prd Typ", etError: Exit Sub
'  Case Is = "WS", "GS", "GP"
'    If ADC("wCpPpSetTypFr") = "" Or ADC("wCpPpSetTypTo") = "" Then _
'       DispMsg "Enter Proper Source/ Target Setting Type", etError: Exit Sub
'    If ADC("wCpPpPrdCtgFr") <> "" Or ADC("wCpPpPrdCtgTo") <> "" Then _
'       DispMsg "Source/ Target Prd Ctg Should Be Blank For This Prd Typ", etError: Exit Sub
'  End Select
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  Call CpyPrdPts
  Call CmdCP_Click
End Sub
Private Sub CpyPrdPts()
  '*** Copy Option will copy the Prd Points from one PrdCtg to another PrdCtg (in case of 'GN' and 'WP')
      'and from one Setting Type to another (in case of 'WS', 'GS' and 'GP')
      'for the given PrdTyp and for the given Design Code, Loc range
  '*** Pass 'C' as the parameter for the CmpString procedure. This will set the values for the
      'CmpStr option of the key frame fields as blank and set proper values for the CmpStr
      'option of the fields in the copy frame
  '*** Cannot copy if Record exists for the given PpPrdCtg, PpSetTyp, PpDmCd, PpLoc range option

  Dim ws_PpCnd As String
  Call CmpString("C")     ' to set the Compare String Property of the ADC controls when the Copy Option Is clicked
  
  ws_PpCnd = IIF(ws_PpCnd <> "", " and " + ws_PpCnd, "")
  ws_PpCnd = IIF(ws_PpCnd <> "", " and ", "") + " (PpDmTcTyp='' or PpDmTcTyp='" + ms_DmTcTyp + "')"
  
  '3.11.2
  If ADC("wCpPpLocFr") = "" Or ADC("wCpPpLocTo") = "" Or (ms_MenuCd <> "PRCLMST" And (ADC("wCpPpSubPrcFr") = "" Or _
     ADC("wCpPpSubPrcTo") = "")) Then DispMsg "All Fields In The Scope Frame Should Be Filled Up", etError: Exit Sub
  
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  'If moCn.RecSeek("Select 'x' from PrdPts where PpTyp='" + ADC("wPpTyp") + "' and " + _
                  "PpPrdCtg= '" + ADC("wCpPpPrdCtgTo") + "' and " + _
                  "PpSetTyp= '" + ADC("wCpPpSetTypTo") + "' " + ws_PpCnd) Then
   ' DispMsg "Cannot Copy as Prd Points Exist For Given PrdCtg/ SetTyp And Given Ranges", etError
    'Exit Sub
  'End If
  
  'If moCn.RecSeek("Select 'x' from PrdPts where PpTyp= '" + ADC("wPpTyp") + "' and " + _
                  "PpLoc= '" + ADC("wCpPpLocTo") + "' " + _
                  IIf(ADC("wCpPpSubPrcFr") = "", "", " and PpSubPrc>= '" + ADC("wCpPpSubPrcFr") + "' ") + _
                  IIf(ADC("wCpPpSubPrcTo") = "", "", " and PpSubPrc<= '" + ADC("wCpPpSubPrcTo") + "' ")) Then
    
  If moCn.RecSeek("Select 'x' from PrdPts where PpTyp= '" + ADC("wPpTyp") + "' and " + _
                  "PpLoc= '" + ADC("wCpPpLocTo") + "' and PpSubPrc= '" + ADC("wCpPpSubPrcTo") + "' and PpPrtKey='" + ctCurrPrtn + "' ") Then
    DispMsg "Cannot Copy as Prd Points Exist For Given Prc/ SubPrc Scope", etError
    Exit Sub
  End If
  '*** Beginning the copy process
  GrdPP.StartCopy
  MWLib.BeginProcess Me, "Copying Production Points ..."
  Dim wRsPrdPts As MwfLib.MDORowSet, wi_Row As Integer
  'Set wRsPrdPts = moCn.OpenRes("Select * from PrdPts where PpTyp='" + ADC("wPpTyp") + "' and " + _
                  "PpPrdCtg= '" + ADC("wCpPpPrdCtgFr") + "' and " + _
                  "PpSetTyp= '" + ADC("wCpPpSetTypFr") + "' " + ws_PpCnd + _
                  " Order By PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts ")
  
  Set wRsPrdPts = moCn.OpenRes("Select * from PrdPts where PpTyp= '" + ADC("wPpTyp") + "' and " + _
                  "PpLoc= '" + ADC("wCpPpLocFr") + "' and " + _
                  "PpSubPrc= '" + ADC("wCpPpSubPrcFr") + "' and PpPrtkey='" + ctCurrPrtn + "' AND " + _
                  " (PpDmTcTyp='" + ms_DmTcTyp + "')" + _
                  " Order By PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmTcTyp,PpDmCd, PpPts ")
  
  ''Set wRsPrdPts = moCn.OpenRes("Select * from PrdPts where PpTyp= '" + ADC("wPpTyp") + "' " + _
                  ws_PpCnd + " Order By PpLoc, PpSubPrc, PpPrdCtg, PpSetTyp, PpDmCd, PpPts ")
  
  With wRsPrdPts
    Do While Not (.EOF Or .BOF)
      MWLib.RunningProcess "Copying Production Points ..."
      GrdPP.AddItem
      wi_Row = GrdPP.Rows - 1
      GrdPP.Value(wi_Row, "PpPrtKey") = ctCurrPrtn    '****** Sachin 3.02.0
      GrdPP.Value(wi_Row, "PpPrdCtg") = !PpPrdCtg
      GrdPP.Value(wi_Row, "PpSetTyp") = !PpSetTyp
      GrdPP.Value(wi_Row, "PpLoc") = ADC("wCpPpLocTo")
      GrdPP.Value(wi_Row, "PpSubPrc") = ADC("wCpPpSubPrcTo")
      GrdPP.Value(wi_Row, "PpDmTcTyp") = !PpDmTcTyp
      GrdPP.Value(wi_Row, "PpDmCd") = !PpDmCd
      GrdPP.Value(wi_Row, "PpPts") = !PpPts
      ' **** Emr 211 **** '
      GrdPP.Value(wi_Row, "PpSkl") = !PpSkl
      GrdPP.Value(wi_Row, "PpRmCtg") = !PpRmCtg
      GrdPP.Value(wi_Row, "PpRmSCtg") = !PpRmSCtg
      GrdPP.Value(wi_Row, "PpFrRmWt") = !PpFrRmWt
      GrdPP.Value(wi_Row, "PpToRmWt") = !PpToRmWt
      ' **** Emr 211 **** '
      GrdPP.SaveRec (wi_Row)
      .MoveNext
      
      'GrdPP.Value(wi_Row, "PpPrdCtg") = ADC("wCpPpPrdCtgTo")
      'GrdPP.Value(wi_Row, "PpSetTyp") = ADC("wCpPpSetTypTo")
      'GrdPP.Value(wi_Row, "PpLoc") = !PpLoc
      'GrdPP.Value(wi_Row, "PpSubPrc") = !PpSubPrc
      'GrdPP.Value(wi_Row, "PpDmCd") = !PpDmCd
      'GrdPP.Value(wi_Row, "PpPts") = !PpPts
      'GrdPP.SaveRec (wi_Row)
    Loop
  End With
  '***** Changed By Shilpa Emr 206 29/10/03 ******
  MWLib.EndProcess Me
  GrdPP.EndCopy
  Set wRsPrdPts = Nothing
  '*** End of the copy process
End Sub
Private Function ColSty(ByVal mColSty As String) As en_fgColStyle
  '*** Sets the Col Style property for the fields in the Grid
  Select Case mColSty
    Case Is = "N"
      ColSty = en_fgColStyle.fgcNormal
    Case Is = "I"
      ColSty = en_fgColStyle.fgcinvisible
    Case Is = "R"
      ColSty = en_fgColStyle.fgcReadOnly
  End Select
End Function
Private Function DataTyp(ByVal mDataTyp As String) As en_FldDataType
  '*** Sets the datatype property for the fields in the Grid
  Select Case mDataTyp
    Case Is = "C"
      DataTyp = fdtCharType
    Case Is = "N"
      DataTyp = fdtNumber
    Case Is = "F"
      DataTyp = fdtFloat
    Case Is = "D"
      DataTyp = fdtDateType
    Case Is = "U"
      DataTyp = fdtConvertToUpper
    Case Is = "T"
      DataTyp = fdtTimeType
  End Select
End Function
Private Sub GrdPropSet()
  '*** Sets all col Properties for the grid based on the TcStr table
  Dim wo_RsTc As MwfLib.MDORowSet, wi_SpaceLen As Single
  
  Set wo_RsTc = moCn.OpenRes("Select * from TcStr where TsTbl='PrdPts' and " + _
                             "TsTyp='" + ADC("wPpTyp") + "' ")
                             
  With GrdPP
    Do While Not (wo_RsTc.EOF Or wo_RsTc.BOF)
      .ColProp(wo_RsTc!TsIdName).Style = ColSty(wo_RsTc!TsStyle)
      
      wi_SpaceLen = (wo_RsTc!TsMaxLen * 2.5) - (Len(wo_RsTc!TsLbl))
      .ColProp(wo_RsTc!TsIdName).Heading = wo_RsTc!TsLbl + IIF(wi_SpaceLen > 0, Space(Abs(wi_SpaceLen)), "")
      
      .ColProp(wo_RsTc!TsIdName).DataType = DataTyp(wo_RsTc!TsDataTyp)
      .ColProp(wo_RsTc!TsIdName).Mask = wo_RsTc!TsMask
      .ColProp(wo_RsTc!TsIdName).MaxLength = wo_RsTc!TsMaxLen
      wo_RsTc.MoveNext
    Loop
  End With
  GrdPP.SetColWidthToHd
  Set wo_RsTc = Nothing
End Sub
Private Sub CmpString(ByVal mTyp As String)
'*** Sets the CmpStr properties for the fields in the Key Frame and the Copy Frame

  '****** Seting values for CmpStr of 'K'ey fields, for Adc.RepCnd (Making Copy fields CmpStr = '')
  If mTyp = "K" Then
    ADC("wPpTyp").CmpStr = "PpTyp= "
    ADC("wPpPrdCtgFr").CmpStr = "PpPrdCtg>= "
    ADC("wPpPrdCtgTo").CmpStr = "PpPrdCtg<= "
    ADC("wPpSetTypFr").CmpStr = "PpSetTyp>= "
    ADC("wPpSetTypTo").CmpStr = "PpSetTyp<= "
    ADC("wPpDmCdFr").CmpStr = "PpDmCd>= "
    ADC("wPpDmCdTo").CmpStr = "PpDmCd<= "
    ADC("wPpLocFr").CmpStr = "PpLoc>= "
    ADC("wPpLocTo").CmpStr = "PpLoc<= "
    'wCpPpPrdCtgFr  wCpPpPrdCtgTo  wCpPpSetTypFr  wCpPpSetTypTo
    ADC("wCpPpLocFr").CmpStr = ""
    'ADC("wCpPpLocTo").CmpStr = ""
    '***** Changed By Shilpa Emr 206 29/10/03 ******
    ADC("wCpPpSubPrcFr").CmpStr = ""
    'ADC("wCpPpSubPrcTo").CmpStr = ""
    'ADC("wCpPpDmCdFr").CmpStr = ""
    'ADC("wCpPpDmCdTo").CmpStr = ""
    '***** Changed By Shilpa Emr 206 29/10/03 ******
  '****** Seting values for CmpStr of 'C'opy fields, for Adc.RepCnd (Making Key fields CmpStr = '')
  ElseIf mTyp = "C" Then
    ADC("wPpTyp").CmpStr = "PpTyp= "
    ADC("wPpPrdCtgFr").CmpStr = ""
    ADC("wPpPrdCtgTo").CmpStr = ""
    ADC("wPpSetTypFr").CmpStr = ""
    ADC("wPpSetTypTo").CmpStr = ""
    ADC("wPpDmCdFr").CmpStr = ""
    ADC("wPpDmCdTo").CmpStr = ""
    ADC("wPpLocFr").CmpStr = ""
    ADC("wPpLocTo").CmpStr = ""
    ADC("wCpPpLocFr").CmpStr = "PpLoc = "
    'ADC("wCpPpLocTo").CmpStr = "PpLoc<= "
    '***** Changed By Shilpa Emr 206 29/10/03 ******
   ' ADC("wCpPpDmCdFr").CmpStr = "PpDmCd>= "
   ' ADC("wCpPpDmCdTo").CmpStr = "PpDmCd<= "
    ADC("wCpPpSubPrcFr").CmpStr = "PpSubPrc = "
    'ADC("wCpPpSubPrcTo").CmpStr = "PpSubPrc <= "
    '***** Changed By Shilpa Emr 206 29/10/03 ******
  End If
End Sub
Private Sub DispFra(ByVal pv_PpFra As en_PpFra)
  '*** This is the common procedure which is used to toggle between hiding and displaying
      'the frames (i.e. if the frame is visible then make it invisible and viseversa)
      'associated with a command button
  
  Select Case pv_PpFra
  Case Is = Cpy
    If FraCP.Visible = True Then
      FraCP.Visible = False
      FraCP.Enabled = False
      Call EnaDisaCmds(False, CmdCP)
      CmdCP.SetFocus
    Else
      FraCP.Visible = True
      FraCP.Enabled = True
      ADC("wCpPpLocFr").SetFocus
      FraCP.ZOrder
      Call EnaDisaCmds(True, CmdCP)
    End If
  End Select
End Sub
Private Sub EnaDisaCmds(ByVal pv_ShowFra As Boolean, Optional ByVal pv_CmdBut As MWCTL_BTN1)
  '*** This procedure is used to give the toggle option (enable and disable) to
      'the command buttons
  
  If Not pv_CmdBut Is Nothing Then
    If pv_ShowFra = True Then
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = vbWhite
'      pv_CmdBut.FontBold = True
      '*** Jenny (Old Code Bef Color)
      pv_CmdBut.CausesValidation = True
    Else
      '*** Jenny (Old Code Bef Color)
'      pv_CmdBut.BackColor = &H8000000F
'      pv_CmdBut.FontBold = False
      '*** Jenny (Old Code Bef Color)
      'pv_CmdBut.CausesValidation = False
    End If
  End If
  FraNKeyAll.Enabled = Not pv_ShowFra
  ADC.AllowSave = Not pv_ShowFra
  CmdCP.Enabled = Not pv_ShowFra
  If Not pv_CmdBut Is Nothing Then pv_CmdBut.Enabled = True
End Sub
Private Sub HideAllFras()
  '*** This procedure hides all the frames other than FraNKeyAll
  FraNKeyAll.Visible = True
  FraCP.Visible = False
  FraCP.Enabled = False
End Sub


'''  ***** Std Code not to be changed **********************************
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
  ws_BtnStr = "CmdCpGo"
  Call SetObjColors(Me, ws_BtnStr)
  '*** Jenny Colour
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
  Cancel = ATXT(Index).Validate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then Cancel = True: Exit Sub
    
    Set moCn = Nothing    '*** (09/08/05)  '*** (Bef speed) bef 24/11/06
End Sub
Private Sub GrdPp_Validate(Cancel As Boolean)
  Cancel = GrdPP.Validate
End Sub
